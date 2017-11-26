' ########################################################################################
' Microsoft Windows
' File: CWebCtx.inc
' Contents: WebBrowser control class
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2017 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "Afx/CAxHost/CAxHost.inc"

NAMESPACE Afx

' ========================================================================================
' Macro for debug
' To allow debugging, define _CWBX_DEBUG_ 1 in your application before including this file.
' ========================================================================================
#ifndef _CWBX_DEBUG_
   #define _CWBX_DEBUG_ 0
#endif
#ifndef _CWBX_DP_
   #define _CWBX_DP_ 1
   #MACRO CWBX_DP(st)
      #if (_CWBX_DEBUG_ = 1)
         OutputDebugStringW(st)
      #endif
   #ENDMACRO
#endif
' ========================================================================================

' // Forward references
DECLARE FUNCTION AfxGetBrowserPtr (BYVAL hwnd AS HWND, BYVAL cID AS WORD) AS Afx_IWebBrowser2 PTR
DECLARE FUNCTION AfxWriteHtml (BYVAL pWebBrowser AS Afx_IWebBrowser2 PTR, BYREF cwsHtml AS CWSTR, BYVAL cr AS BOOLEAN = FALSE) AS HRESULT
DECLARE FUNCTION AfxGetActiveElementId (BYVAL pWebBrowser AS Afx_IWebBrowser2 PTR) AS CWSTR
DECLARE FUNCTION AfxGetElementValueById (BYVAL pWebBrowser AS Afx_IWebBrowser2 PTR, BYREF cwsId AS CWSTR) AS VARIANT
DECLARE FUNCTION AfxSetElementValueById (BYVAL pWebBrowser AS Afx_IWebBrowser2 PTR, BYREF cwsId AS CWSTR, BYVAL vValue AS VARIANT) AS HRESULT
DECLARE FUNCTION AfxGetElementInnerHtmlById (BYVAL pWebBrowser AS Afx_IWebBrowser2 PTR, BYREF cwsId AS CWSTR) AS CWSTR
DECLARE FUNCTION AfxSetElementInnerHtmlById (BYVAL pWebBrowser AS Afx_IWebBrowser2 PTR, BYREF cwsId AS CWSTR, BYREF cwsHtml AS CWSTR) AS HRESULT
DECLARE FUNCTION AfxSetElementFocusById (BYVAL pWebBrowser AS Afx_IWebBrowser2 PTR, BYREF cwsId AS CWSTR) AS HRESULT
DECLARE FUNCTION AfxGetBodyInnerText (BYVAL pWebBrowser AS Afx_IWebBrowser2 PTR) AS CWSTR
DECLARE FUNCTION AfxGetBodyInnerHtml (BYVAL pWebBrowser AS Afx_IWebBrowser2 PTR) AS CWSTR

TYPE CWebBrowserEvents_ AS CWebBrowserEvents
TYPE CDocHostUIHandler2_ AS CDocHostUIHandler2

' ========================================================================================
' CWebCtx class
' ========================================================================================
TYPE CWebCtx

   m_pAxHost AS CAxHost PTR                         ' // Pointer to the OLE container
   m_pWebBrowser AS Afx_IWebBrowser2 PTR            ' // Pointer to the IWebBrowser interface
   m_pEvtObj AS CWebBrowserEvents_ PTR              ' // Pointer to the implemented Web Browser events class
   m_pDocHostUIHandler AS CDocHostUIHandler2_ PTR   ' // Pointer to the implemented DocHostUIHandler interface

   DECLARE CONSTRUCTOR (BYVAL pWindow AS CWindow PTR, BYVAL cID AS LONG_PTR, _
      BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
      BYVAL dwStyle AS DWORD = WS_VISIBLE OR WS_CHILD OR WS_TABSTOP, BYVAL dwExStyle AS DWORD = 0, BYVAL pAmbientDisp AS CAXHOST_AMBIENTDISP PTR = NULL)
   DECLARE DESTRUCTOR
   DECLARE FUNCTION AxHostPtr () AS CAxHost PTR
   DECLARE FUNCTION BrowserPtr () AS Afx_IWebBrowser2 PTR
   DECLARE FUNCTION hWindow () AS HWND
   DECLARE FUNCTION CtrlID () AS LONG
   DECLARE FUNCTION Navigate (BYVAL pwszUrl AS WSTRING PTR, BYVAL Flags AS VARIANT PTR = NULL, BYVAL TargetFrameName AS VARIANT PTR = NULL, BYVAL PostData AS VARIANT PTR = NULL, BYVAL Headers AS VARIANT PTR = NULL) AS HRESULT
   DECLARE FUNCTION Navigate (BYVAL vUrl AS VARIANT PTR, BYVAL Flags AS VARIANT PTR = NULL, BYVAL TargetFrameName AS VARIANT PTR = NULL, BYVAL PostData AS VARIANT PTR = NULL, BYVAL Headers AS VARIANT PTR = NULL) AS HRESULT
   DECLARE FUNCTION WaitForPageLoad (BYVAL dbTimeout AS DOUBLE = 10) AS tagREADYSTATE
   DECLARE FUNCTION Advise () AS HRESULT
   DECLARE FUNCTION Unadvise () AS HRESULT
   DECLARE FUNCTION SetUIHandler () AS HRESULT
   DECLARE FUNCTION SetEventProc (BYVAL pwszEventName AS WSTRING PTR, BYVAL pProc AS ANY PTR) AS HRESULT
   DECLARE FUNCTION SetUIEventProc (BYVAL pwszEventName AS WSTRING PTR, BYVAL pProc AS ANY PTR) AS HRESULT

   ' // IWebBrowser2 interface
   DECLARE PROPERTY Silent (BYVAL pbSilent AS VARIANT_BOOL)
   DECLARE PROPERTY Silent () AS VARIANT_BOOL
'   DECLARE PROPERTY Resizable (BYVAL bResize AS VARIANT_BOOL)
'   DECLARE PROPERTY Resizable () AS VARIANT_BOOL
   DECLARE FUNCTION GoBack () AS HRESULT
   DECLARE FUNCTION GoForward () AS HRESULT
   DECLARE FUNCTION GoHome () AS HRESULT
   DECLARE FUNCTION GoSearch () AS HRESULT
   DECLARE FUNCTION Refresh () AS HRESULT
   DECLARE FUNCTION Refresh2 (BYVAL nLevel AS RefreshConstants) AS HRESULT
   DECLARE FUNCTION Stop () AS HRESULT
   DECLARE PROPERTY Document () AS IHtmlDocument2 PTR
   DECLARE PROPERTY LocationName () AS CWSTR
   DECLARE PROPERTY LocationURL () AS CWSTR
   DECLARE PROPERTY Busy () AS BOOLEAN
   DECLARE FUNCTION QueryStatusWB (BYVAL cmdID AS OLECMDID) AS OLECMDF
   DECLARE FUNCTION ExecWB (BYVAL cmdID AS OLECMDID, BYVAL cmdexecopt AS OLECMDEXECOPT, BYVAL pvaIn AS VARIANT PTR = NULL, BYVAL pvaOut AS VARIANT PTR = NULL) AS HRESULT
   DECLARE PROPERTY ReadyState () AS tagREADYSTATE
   DECLARE PROPERTY RegisterAsBrowser () AS BOOLEAN
   DECLARE PROPERTY RegisterAsBrowser (BYVAL bRegister AS BOOLEAN)
   DECLARE PROPERTY RegisterAsDropTarget () AS BOOLEAN
   DECLARE PROPERTY RegisterAsDropTarget (BYVAL bRegister AS BOOLEAN)
   DECLARE FUNCTION PrintPreview () AS HRESULT
   DECLARE FUNCTION PrintPage (BYVAL bPromptUser AS BOOLEAN = TRUE) AS HRESULT
   DECLARE FUNCTION PageSetUp () AS HRESULT
   DECLARE FUNCTION PageProperties () AS HRESULT
   DECLARE FUNCTION SaveAs () AS HRESULT
   DECLARE FUNCTION Find () AS HRESULT
   DECLARE FUNCTION ShowSource () AS HRESULT
   DECLARE FUNCTION InternetOptions () AS HRESULT

   ' // IHtmlDoc2 wrappers
   DECLARE FUNCTION SetFocus () AS HRESULT
   DECLARE FUNCTION WriteHtml (BYREF cwsHtml AS CWSTR, BYVAL cr AS BOOLEAN = FALSE) AS HRESULT
   DECLARE FUNCTION GetBodyInnerText () AS CWSTR
   DECLARE FUNCTION GetBodyInnerHtml () AS CWSTR
   DECLARE FUNCTION GetActiveElementId () AS CWSTR
   DECLARE FUNCTION GetElementValueById (BYREF cwsId AS CWSTR) AS VARIANT
   DECLARE FUNCTION SetElementValueById (BYREF cwsId AS CWSTR, BYVAL vValue AS VARIANT) AS HRESULT
   DECLARE FUNCTION GetElementInnerHtmlById (BYREF cwsId AS CWSTR) AS CWSTR
   DECLARE FUNCTION SetElementInnerHtmlById (BYREF cwsId AS CWSTR, BYREF cwsHtml AS CWSTR) AS HRESULT
   DECLARE FUNCTION SetElementFocusById (BYREF cwsId AS CWSTR) AS HRESULT

END TYPE
' ========================================================================================

' // DocHostUIHandler events class
#include once "Afx/CAxHost/CDocHostUIHandler.inc"
' // HtmlDocumentEvents2 class
#include once "Afx/CAxHost/CHtmlDocumentEvents.inc"
' // WebBrowser events class
#include once "Afx/CAxHost/CWebBrowserEvents.inc"

' ========================================================================================
' Constructor
' ========================================================================================
CONSTRUCTOR CWebCtx (BYVAL pWindow AS CWindow PTR, BYVAL cID AS LONG_PTR, _
BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
BYVAL dwStyle AS DWORD = WS_VISIBLE OR WS_CHILD OR WS_TABSTOP, BYVAL dwExStyle AS DWORD = 0, BYVAL pAmbientDisp AS CAXHOST_AMBIENTDISP PTR = NULL)
   CWBX_DP("CONSTRUCTOR CWebCtx")

   ' // Create an instance of the OLE container
   m_pAxHost = NEW CAxHost(pWindow, cID, "Shell.Explorer", x, y, nWidth, nHeight, dwStyle, dwExStyle, pAmbientDisp)
   ' // Get a reference to the IWebBrowser2 interface
   IF m_pAxHost THEN m_pWebBrowser = CAST(Afx_IWebBrowser2 PTR, CAST(ULONG_PTR, m_pAxHost->OcxDispPtr))
   ' // Navigate to a blank page to allow customization by calling the SetUIHandler method.
   ' // Otherwise, it will fail when trying to get the IHtmlDocument2 interface.
   IF m_pWebBrowser THEN
      DIM vUrl AS VARIANT : vUrl.vt = VT_BSTR : vUrl.bstrVal = SysAllocString("about:blank")
      m_pWebBrowser->Navigate2(@vUrl, NULL, NULL, NULL, NULL)
      VariantClear @vurl
      ' // Set the IDocHostUIHandler interface
      this.SetUIHandler
      ' // Connect events
      this.Advise
   END IF

END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destructor
' ========================================================================================
DESTRUCTOR CWebCtx
   CWBX_DP("BEGIN DESTRUCTOR CWebCtx")
   ' // Stop navigation
   IF m_pWebBrowser THEN m_pWebBrowser->Stop
   ' // Disconnect events and delete the OLE container, that in turn
   ' // releases the IWebBrowser2 interface
   IF m_pAxHost THEN
      m_pAxHost->Unadvise
      DELETE m_pAxHost
   END IF
   ' // Delete the WebBrowser event's class
   ' // Important: Don't delete it before deleting the OLE container because
   ' // it deletes the IDocHostUIHandler class and you will get a GPF when the
   ' // OLE container will call IDocHostUIHandler to send a message.
   IF m_pEvtObj THEN DELETE m_pEvtObj
   IF m_pDocHostUIHandler THEN DELETE m_pDocHostUIHandler
   CWBX_DP("END DESTRUCTOR CWebCtx")
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns a pointer to the OLE container class.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.AxHostPtr () AS CAxHost PTR
   CWBX_DP("CWebCtx.AxHostPtr")
   RETURN m_pAxHost
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a direct pointer to the IWebBrowser2 interface of the hosted WebBrowser control.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.BrowserPtr () AS Afx_IWebBrowser2 PTR
   CWBX_DP("CWebCtx.BrowserPtr")
   RETURN m_pWebBrowser
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the handle of the OLE container hosting window
' ========================================================================================
PRIVATE FUNCTION CWebCtx.hWindow () AS HWND
   CWBX_DP("CWebCtx.hWindow")
   RETURN m_pAxHost->hWindow
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the identifier of the OLE container hosting window
' ========================================================================================
PRIVATE FUNCTION CWebCtx.CtrlID () AS LONG
   CWBX_DP("CWebCtx.CtrlID")
   RETURN GetDlgCtrlID(m_pAxHost->hWindow)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Navigates to the specified URL.
' Note: To clear a web page, simply navigate to "about:blank".
' ========================================================================================
PRIVATE FUNCTION CWebCtx.Navigate (BYVAL pwszUrl AS WSTRING PTR, BYVAL Flags AS VARIANT PTR = NULL, BYVAL TargetFrameName AS VARIANT PTR = NULL, BYVAL PostData AS VARIANT PTR = NULL, BYVAL Headers AS VARIANT PTR = NULL) AS HRESULT
   CWBX_DP("CWebCtx.Navigate - pwszUrl")
   DIM vUrl AS VARIANT : vUrl.vt = VT_BSTR : vUrl.bstrVal = SysAllocString(pwszUrl)
   IF m_pWebBrowser THEN FUNCTION = m_pWebBrowser->Navigate2(@vUrl, Flags, TargetFrameName, PostData, Headers)
   VariantClear(@vUrl)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CWebCtx.Navigate (BYVAL vUrl AS VARIANT PTR, BYVAL Flags AS VARIANT PTR = NULL, BYVAL TargetFrameName AS VARIANT PTR = NULL, BYVAL PostData AS VARIANT PTR = NULL, BYVAL Headers AS VARIANT PTR = NULL) AS HRESULT
   CWBX_DP("CWebCtx.Navigate - vUrl")
   IF m_pWebBrowser THEN FUNCTION = m_pWebBrowser->Navigate2(vUrl, Flags, TargetFrameName, PostData, Headers)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Waits until the page had been fully downloaded or te timeout has expired.
' - dbTimeout: Timeout in seconds.
' Return value:
' READYSTATE_UNINITIALIZED = 0
'   Default initialization state.
' READYSTATE_LOADING = 1
'   Object is currently loading its properties.
' READYSTATE_LOADED = 2
'   Object has been initialized.
' READYSTATE_INTERACTIVE = 3
'   Object is interactive, but not all of its data is available.
' READYSTATE_COMPLETE = 4
'   Object has received all of its data.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.WaitForPageLoad (BYVAL dbTimeout AS DOUBLE = 10) AS tagREADYSTATE
   CWBX_DP("CWebCtx.WaitForPageLoad")
   DIM lReadyState AS tagREADYSTATE
   DIM starttime AS DOUBLE = timer
   DO
      ' // Processes pending messages
      ' // Needed to allow the page to load
      AfxPumpMessages
      ' // Retrieves the ready state
      m_pWebBrowser->get_ReadyState(@lReadyState)
      IF lReadyState = READYSTATE_COMPLETE THEN EXIT DO
      IF timer - starttime > dbTimeout THEN EXIT DO
      ' // Needed to update timer
      SLEEP 1, 1
   LOOP
   RETURN lReadyState
END FUNCTION
' ========================================================================================

' ========================================================================================
' Establishes a connection between the connection point object and the client's sink.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.Advise () AS HRESULT
   CWBX_DP("CWebCtx.Advise")
   ' // Create an instance of our CWebBrowserEvents class and connect the WebBrowser events
   m_pEvtObj = NEW CWebBrowserEvents(@this)
   m_pEvtObj->m_pWebBrowser = m_pWebBrowser
   RETURN m_pAxHost->Advise(CAST(IDispatch PTR, CAST(ANY PTR, m_pEvtObj)), CAST(IID PTR, @DIID_DWebBrowserEvents2))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Releases the events connection.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.Unadvise () AS HRESULT
   CWBX_DP("CWebCtx.Unadvise")
   RETURN m_pAxHost->Unadvise
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets our implementation of the IDocHostUIHandler interface to customize the WebBrowser.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.SetUIHandler () AS HRESULT
   CWBX_DP("-BEGIN CWebCtx.SetUIHandler")

   ' // If it is already set, exit
   IF m_pDocHostUIHandler THEN RETURN S_OK

   ' // After the MSHTML document has been loaded, we retrieve a reference to its
   ' // ICustomDoc interface and give him a pointer to our IDocHostUIHandler interface
   ' // to allow for customization.
   IF m_pWebBrowser = NULL THEN RETURN E_POINTER

   ' // Get a reference to the active document
   DIM pIHTMLDocument2 AS IHTMLDocument2 PTR
   m_pWebBrowser->get_Document(@cast(ANY PTR, pIHTMLDocument2))
   IF pIHTMLDocument2 = NULL THEN EXIT FUNCTION

   ' // Get a reference to the CustomDoc interface
   DIM pICustomDoc AS ICustomDoc PTR
   ' extern IID_ICustomDoc is broken and causes the linker to fail
   ' IID_ICustomDoc = "{3050F3F0-98B5-11CF-BB82-00AA00BDCE0B}"
   DIM IID_ICustomDoc_ AS GUID = (&h3050F3F0, &h98B5, &h11CF, {&hBB, &h82, &h00, &hAA, &h00, &hBD, &hCE, &h0B})
   pIHTMLDocument2->lpvtbl->QueryInterface(pIHTMLDocument2, @IID_ICustomDoc_, @pICustomDoc)
   IUnknown_Release(pIHTMLDocument2)
   IF pICustomDoc = NULL THEN EXIT FUNCTION

   ' // Create an instance of the IDocHostUIHandler class
   IF m_pDocHostUIHandler = NULL THEN m_pDocHostUIHandler = NEW CDocHostUIHandler2(@this, m_pWebBrowser)
   ' // Set our IDocHostUIHandler interface for MSHTML
   ' // MSHTML will release its previous IDocHostUIHandler interface
   ' // (if one is present) and call pDocHostUIHandler's AddRef method.
   IF m_pDocHostUIHandler THEN pICustomDoc->lpvtbl->SetUIHandler(pICustomDoc, _
      CAST(IDocHostUIHandler PTR, CAST(ULONG_PTR, m_pDocHostUIHandler)))
   AfxSafeRelease(pICustomDoc)

   CWBX_DP("-END CWebCtx.SetUIHandler")
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets pointers to user implemented callback procedures.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.SetEventProc (BYVAL pwszEventName AS WSTRING PTR, BYVAL pProc AS ANY PTR) AS HRESULT
   CWBX_DP("CWebCtx.SetEventHandler")
   IF pwszEventName = NULL THEN RETURN E_INVALIDARG
   SELECT CASE LCASE(*pwszEventName)
      CASE "statustextchange" : IF m_pEvtObj THEN m_pEvtObj->m_StatusTextChangeProc = pProc
      CASE "downloadcomplete" : IF m_pEvtObj THEN m_pEvtObj->m_DownloadCompleteProc = pProc
      CASE "commandstatechange" : IF m_pEvtObj THEN m_pEvtObj->m_CommandStateChangeProc = pProc
      CASE "downloadbegin" : IF m_pEvtObj THEN m_pEvtObj->m_DownloadBeginProc = pProc
      CASE "progresschange" : IF m_pEvtObj THEN m_pEvtObj->m_ProgressChangeProc = pProc
      CASE "propertychange" : IF m_pEvtObj THEN m_pEvtObj->m_PropertyChangeProc = pProc
      CASE "titlechange" : IF m_pEvtObj THEN m_pEvtObj->m_TitleChangeProc = pProc
      CASE "printtemplateinstantiation" : IF m_pEvtObj THEN m_pEvtObj->m_PrintTemplateInstantiationProc = pProc
      CASE "printtemplateteardown" : IF m_pEvtObj THEN m_pEvtObj->m_PrintTemplateTeardownProc = pProc
      CASE "beforenavigate2" : IF m_pEvtObj THEN m_pEvtObj->m_BeforeNavigate2Proc = pProc
      CASE "navigatecomplete2" : IF m_pEvtObj THEN m_pEvtObj->m_NavigateComplete2Proc = pProc
      CASE "onvisible" : IF m_pEvtObj THEN m_pEvtObj->m_OnVisibleProc = pProc
      CASE "documentcomplete" : IF m_pEvtObj THEN m_pEvtObj->m_DocumentCompleteProc = pProc
      CASE "windowsetresizable" : IF m_pEvtObj THEN m_pEvtObj->m_WindowSetResizableProc = pProc
      CASE "windowclosing" : IF m_pEvtObj THEN m_pEvtObj->m_WindowClosingProc = pProc
      CASE "windowsetleft" : IF m_pEvtObj THEN m_pEvtObj->m_WindowSetLeftProc = pProc
      CASE "windowsettop" : IF m_pEvtObj THEN m_pEvtObj->m_WindowSetTopProc = pProc
      CASE "windowsetwidth" : IF m_pEvtObj THEN m_pEvtObj->m_WindowSetWidthProc = pProc
      CASE "windowsetheight" : IF m_pEvtObj THEN m_pEvtObj->m_WindowSetHeightProc = pProc
      CASE "clienttohostwindow" : IF m_pEvtObj THEN m_pEvtObj->m_ClientToHostWindowProc = pProc
      CASE "setsecurelockicon" : IF m_pEvtObj THEN m_pEvtObj->m_SetSecureLockIconProc = pProc
      CASE "filedownload" : IF m_pEvtObj THEN m_pEvtObj->m_FileDownloadProc = pProc
      CASE "navigateerror" : IF m_pEvtObj THEN m_pEvtObj->m_NavigateErrorProc = pProc
      CASE "privacyimpactedstatechange" : IF m_pEvtObj THEN m_pEvtObj->m_PrivacyImpactedStateChangeProc = pProc
      CASE "newwindow2" : IF m_pEvtObj THEN m_pEvtObj->m_NewWindow2Proc = pProc
      CASE "newwindow3" : IF m_pEvtObj THEN m_pEvtObj->m_NewWindow3Proc = pProc
      CASE "windowstatechanged" : IF m_pEvtObj THEN m_pEvtObj->m_WindowStateChangedProc = pProc
      CASE "htmldocumentevents"
            IF m_pEvtObj THEN
               IF m_pEvtObj->m_pHTMLDocumentEvents2 THEN
                  m_pEvtObj->m_pHTMLDocumentEvents2->m_pHtmlDocumentEventsProc = pProc
               END IF
            END IF
      CASE ELSE
         RETURN E_INVALIDARG
   END SELECT
END FUNCTION
' ========================================================================================

' ========================================================================================
' Set pointer to user implemented callback procedures.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.SetUIEventProc (BYVAL pwszEventName AS WSTRING PTR, BYVAL pProc AS ANY PTR) AS HRESULT
   CWBX_DP("CWebCtx.SetUIEventHandler")
   IF pwszEventName = NULL THEN RETURN E_INVALIDARG
   SELECT CASE LCASE(*pwszEventName)
      CASE "showcontextmenu" : IF m_pDocHostUIHandler THEN m_pDocHostUIHandler->m_ShowContextMenuProc = pProc
      CASE "gethostinfo" : IF m_pDocHostUIHandler THEN m_pDocHostUIHandler->m_GetHostInfoProc = pProc
      CASE "showui" : IF m_pDocHostUIHandler THEN m_pDocHostUIHandler->m_ShowUIProc = pProc
      CASE "hideui" : IF m_pDocHostUIHandler THEN m_pDocHostUIHandler->m_HideUIProc = pProc
      CASE "updateui" : IF m_pDocHostUIHandler THEN m_pDocHostUIHandler->m_UpdateUIProc = pProc
      CASE "enablemodeless" : IF m_pDocHostUIHandler THEN m_pDocHostUIHandler->m_EnableModelessProc = pProc
      CASE "ondocwindowactivate" : IF m_pDocHostUIHandler THEN m_pDocHostUIHandler->m_OnDocWindowActivateProc = pProc
      CASE "onframewindowactivate" : IF m_pDocHostUIHandler THEN m_pDocHostUIHandler->m_OnFrameWindowActivateProc = pProc
      CASE "resizeborder" : IF m_pDocHostUIHandler THEN m_pDocHostUIHandler->m_ResizeBorderProc = pProc
      CASE "translateaccelerator" : IF m_pDocHostUIHandler THEN m_pDocHostUIHandler->m_TranslateAcceleratorProc = pProc
      CASE "getoptionkeypath" : IF m_pDocHostUIHandler THEN m_pDocHostUIHandler->m_GetOptionKeyPathProc = pProc
      CASE "getdroptarget" : IF m_pDocHostUIHandler THEN m_pDocHostUIHandler->m_GetDropTargetProc = pProc
      CASE "getexternal" : IF m_pDocHostUIHandler THEN m_pDocHostUIHandler->m_GetExternalProc = pProc
      CASE "translateurl" : IF m_pDocHostUIHandler THEN m_pDocHostUIHandler->m_TranslateUrlProc = pProc
      CASE "filterdataobject" : IF m_pDocHostUIHandler THEN m_pDocHostUIHandler->m_FilterDataObjectProc = pProc
      CASE "getoverridekeypath" : IF m_pDocHostUIHandler THEN m_pDocHostUIHandler->m_GetOverrideKeyPathProc = pProc
      CASE ELSE
         RETURN E_INVALIDARG
   END SELECT
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the focus in the hosted document.
' Return value:
' - S_OK if successful, or an error value otherwise.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.SetFocus () AS HRESULT
   CWBX_DP("CWebCtx.SetFocus")
   IF m_pWebBrowser = NULL THEN RETURN E_POINTER
   DIM pIHTMLDocument2 AS IHTMLDocument2 PTR
   m_pWebBrowser->get_Document(@cast(ANY PTR, pIHTMLDocument2))
   IF pIHTMLDocument2 = NULL THEN RETURN E_NOINTERFACE
   DIM pWindow2 AS IHTMLWindow2 PTR
   pIHTMLDocument2->lpvtbl->get_parentWindow(pIHTMLDocument2, @pWindow2)
   AfxSafeRelease(pIHTMLDocument2)
   IF pWindow2 = NULL THEN RETURN E_NOINTERFACE
   pWindow2->lpvtbl->focus(pWindow2)
   AfxSafeRelease(pWindow2)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes one or more HTML expressions to a document.
' Parameters
' - cbsHtml: Text and HTML tags to write.
' - cr: Write the HTML text followed by a carriage return.
' Remarks
'   In HTML, the carriage return is ignored unless it occurs in preformatted text.
'   Note  When document.IHTMLDocument2::write or document.IHTMLDocument2::writeln is used
'   in an event handler, you must also use document.IHTMLDocument2::close.
' Return value:
' - S_OK if successful, or an error value otherwise.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.WriteHtml (BYREF cwsHtml AS CWSTR, BYVAL cr AS BOOLEAN = FALSE) AS HRESULT
   CWBX_DP("CWebCtx WriteHtml")
   IF m_pWebBrowser = NULL THEN RETURN E_POINTER
   DIM pIHTMLDocument2 AS IHTMLDocument2 PTR
   DIM hr AS HRESULT = m_pWebBrowser->get_Document(@CAST(ANY PTR, pIHTMLDocument2))
   IF hr <> S_OK OR pIHTMLDocument2 = NULL THEN RETURN hr
   ' // Create a safearray of variants of one element
   DIM psarray AS SafeArray PTR
   DIM rgsabounds(0) AS SAFEARRAYBOUND = {(1, 0)}
   psarray = SafeArrayCreate(VT_VARIANT, 1, @rgsabounds(0))
   IF psarray = NULL THEN RETURN E_FAIL
   ' // Fill the safearray with the script
   DIM ix AS LONG = 0
   DIM vHtml AS VARIANT
   vHtml.vt = VT_BSTR
   vHtml.bstrVal = SysAllocString(**cwsHtml)
   SafeArrayPutElement(psarray, @ix, @vHtml)
   ' // Write the string
   IF cr THEN
      hr = pIHTMLDocument2->lpvtbl->writeln(pIHTMLDocument2, psarray)
   ELSE
      hr = pIHTMLDocument2->lpvtbl->write(pIHTMLDocument2, psarray)
   END IF
   ' // Destroy the safearray
   SafeArrayDestroy psarray
   ' // Free the pIHTMLDocument2 interface
   pIHTMLDocument2->lpvtbl->Release(pIHTMLDocument2)
   ' // Clear the variant
   VariantClear @vHtml
   RETURN hr
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a string that represents the text between the start and end body tags without
' any associated HTML.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.GetBodyInnerText () AS CWSTR
   CWBX_DP("CWebCtx GetBodyInnerText")
   IF m_pWebBrowser = NULL THEN RETURN ""
   DIM pIHTMLDocument2 AS IHTMLDocument2 PTR
   DIM hr AS HRESULT = m_pWebBrowser->get_Document(@CAST(ANY PTR, pIHTMLDocument2))
   IF hr <> S_OK OR pIHTMLDocument2 = NULL THEN RETURN ""
   DIM pIHTMLElement AS IHTMLElement PTR
   hr = pIHTMLDocument2->lpvtbl->get_body(pIHTMLDocument2, @pIHTMLElement)
   pIHTMLDocument2->lpvtbl->Release(pIHTMLDocument2)
   IF hr <> S_OK OR pIHTMLElement = NULL THEN RETURN ""
   DIM bstrText AS AFX_BSTR
   hr = pIHTMLElement->lpvtbl->get_innerText(pIHTMLElement, @bstrText)
   pIHTMLElement->lpvtbl->Release(pIHTMLElement)
   DIM cwsText AS CWSTR = *bstrText
   SysFreeString bstrText
   RETURN cwsText
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a string that represents the text and html elements between the start and end body tags.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.GetBodyInnerHtml () AS CWSTR
   CWBX_DP("CWebCtx GetBodyInnerHtml")
   IF m_pWebBrowser = NULL THEN RETURN ""
   DIM pIHTMLDocument2 AS IHTMLDocument2 PTR
   DIM hr AS HRESULT = m_pWebBrowser->get_Document(@CAST(ANY PTR, pIHTMLDocument2))
   IF hr <> S_OK OR pIHTMLDocument2 = NULL THEN RETURN ""
   DIM pIHTMLElement AS IHTMLElement PTR
   hr = pIHTMLDocument2->lpvtbl->get_body(pIHTMLDocument2, @pIHTMLElement)
   pIHTMLDocument2->lpvtbl->Release(pIHTMLDocument2)
   IF hr <> S_OK OR pIHTMLElement = NULL THEN RETURN ""
   DIM bstrText AS AFX_BSTR
   hr = pIHTMLElement->lpvtbl->get_innerHtml(pIHTMLElement, @bstrText)
   pIHTMLElement->lpvtbl->Release(pIHTMLElement)
   DIM cwsText AS CWSTR = *bstrText
   SysFreeString bstrText
   RETURN cwsText
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the ID of the active element (the object that has the focus when the parent
' document has focus).
' Return value: The ID of the active element.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.GetActiveElementId () AS CWSTR
   CWBX_DP("CWebCtx GetActiveElementId")
   IF m_pWebBrowser = NULL THEN RETURN ""
   DIM pIHTMLDocument2 AS IHTMLDocument2 PTR
   DIM hr AS HRESULT = m_pWebBrowser->get_Document(@CAST(ANY PTR, pIHTMLDocument2))
   IF hr <> S_OK OR pIHTMLDocument2 = NULL THEN RETURN ""
   DIM pIHTMLElement AS IHTMLElement PTR
   hr = pIHTMLDocument2->lpvtbl->get_activeElement(pIHTMLDocument2, @pIHTMLElement)
   pIHTMLDocument2->lpvtbl->Release(pIHTMLDocument2)
   IF hr <> S_OK OR pIHTMLElement = NULL THEN RETURN ""
   DIM bstrId AS AFX_BSTR
   hr = pIHTMLElement->lpvtbl->get_id(pIHTMLElement, @bstrId)
   pIHTMLElement->lpvtbl->Release(pIHTMLElement)
   DIM cwsId AS CWSTR = *bstrId
   SysFreeString bstrId
   RETURN cwsId
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the value attribute of the specified identifier.
' Parameter:
' - cbsId = The identifier.
' Return Value:
' - A variant containing the value as defined by the attribute.
' Remarks:
'   This method performs a case insensitive property search.
'   If two or more attributes have the same name (differing only in uppercase and lowercase
'   letters) this function retrieves values only for the last attribute created with this
'   name, and ignores all other attributes with the same name.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.GetElementValueById (BYREF cwsId AS CWSTR) AS VARIANT
   CWBX_DP("CWebCtx GetElementValueById")
   IF m_pWebBrowser = NULL THEN EXIT FUNCTION
   ' // Get a reference to the IHTMLDocument2 interface
   DIM pIHTMLDocument2 AS IHTMLDocument2 PTR
   DIM hr AS HRESULT = m_pWebBrowser->get_Document(@CAST(ANY PTR, pIHTMLDocument2))
   IF hr <> S_OK OR pIHTMLDocument2 = NULL THEN EXIT FUNCTION
   ' // Get a reference to the IHTMLDocument3 interface
   DIM pIHTMLDocument3 AS IHTMLDocument3 PTR
   hr = pIHTMLDocument2->lpvtbl->QueryInterface(pIHTMLDocument2, @IID_IHTMLDocument3, @pIHTMLDocument3)
   pIHTMLDocument2->lpvtbl->Release(pIHTMLDocument2)
   IF hr <> S_OK OR pIHTMLDocument3 = NULL THEN EXIT FUNCTION
   ' // Get a reference to the input element
   DIM pIHTMLElement AS IHTMLElement PTR
   DIM bstrId AS AFX_BSTR = SysAllocString(**cwsId)
   hr = pIHTMLDocument3->lpvtbl->getElementById(pIHTMLDocument3, bstrId, @pIHTMLElement)
   pIHTMLDocument3->lpvtbl->Release(pIHTMLDocument3)
   SysFreeString bstrId
   IF hr <> S_OK OR pIHTMLElement = NULL THEN EXIT FUNCTION
   ' // Get the value
   DIM vValue AS VARIANT, bstrAttr AS AFX_BSTR = SysAllocString("value")
   hr = pIHTMLElement->lpvtbl->getAttribute(pIHTMLElement, bstrAttr, 0, @vValue)
   pIHTMLElement->lpvtbl->Release(pIHTMLElement)
   SysFreeString bstrAttr
   RETURN vValue
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the value attribute of the specified identifier.
' Parameter:
' - cbsId = The identifier.
' - cvValue = Variant that specifies the string, number, or Boolean to assign to the attribute.
' Remarks:
'   This method performs a case insensitive property search.
'   If two or more attributes have the same name (differing only in uppercase and lowercase
'   letters) this function sets values only for the last attribute created with this
'   name, and ignores all other attributes with the same name.
' Return value:
' - S_OK if successful, or an error value otherwise.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.SetElementValueById (BYREF cwsId AS CWSTR, BYVAL vValue AS VARIANT) AS HRESULT
   CWBX_DP("CWebCtx SetElementValueById")
   IF m_pWebBrowser = NULL THEN RETURN E_POINTER
   ' // Get a reference to the IHTMLDocument2 interface
   DIM pIHTMLDocument2 AS IHTMLDocument2 PTR
   DIM hr AS HRESULT = m_pWebBrowser->get_Document(@CAST(ANY PTR, pIHTMLDocument2))
   IF hr <> S_OK OR pIHTMLDocument2 = NULL THEN RETURN hr
   ' // Get a reference to the IHTMLDocument3 interface
   DIM pIHTMLDocument3 AS IHTMLDocument3 PTR
   hr = pIHTMLDocument2->lpvtbl->QueryInterface(pIHTMLDocument2, @IID_IHTMLDocument3, @pIHTMLDocument3)
   pIHTMLDocument2->lpvtbl->Release(pIHTMLDocument2)
   IF hr <> S_OK OR pIHTMLDocument3 = NULL THEN RETURN hr
   ' // Get a reference to the input element
   DIM pIHTMLElement AS IHTMLElement PTR
   DIM bstrId AS AFX_BSTR = SysAllocString(**cwsId)
   hr = pIHTMLDocument3->lpvtbl->getElementById(pIHTMLDocument3, bstrId, @pIHTMLElement)
   pIHTMLDocument3->lpvtbl->Release(pIHTMLDocument3)
   SysFreeString bstrId
   IF hr <> S_OK OR pIHTMLElement = NULL THEN RETURN hr
   ' // Set the value
   DIM bstrAttr AS AFX_BSTR = SysAllocString("value")
   hr = pIHTMLElement->lpvtbl->setAttribute(pIHTMLElement, bstrAttr, vValue, 0)
   pIHTMLElement->lpvtbl->Release(pIHTMLElement)
   SysFreestring bstrAttr
   RETURN hr
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the HTML between the start and end tags of the object.
' Parameter:
' - cbsId = The idenfier.
' Return value:
' - A string containing the HTML text.
' Remarks:
'   This method performs a case insensitive property search.
'   If two or more attributes have the same name (differing only in uppercase and lowercase
'   letters) this function retrieves values only for the last attribute created with this
'   name, and ignores all other attributes with the same name.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.GetElementInnerHtmlById (BYREF cwsId AS CWSTR) AS CWSTR
   CWBX_DP("CWebCtx GetElementInnerHtmlById")
   IF m_pWebBrowser = NULL THEN RETURN ""
   ' // Get a reference to the IHTMLDocument2 interface
   DIM pIHTMLDocument2 AS IHTMLDocument2 PTR
   DIM hr AS HRESULT = m_pWebBrowser->get_Document(@CAST(ANY PTR, pIHTMLDocument2))
   IF hr <> S_OK OR pIHTMLDocument2 = NULL THEN RETURN ""
   ' // Get a reference to the IHTMLDocument3 interface
   DIM pIHTMLDocument3 AS IHTMLDocument3 PTR
   hr = pIHTMLDocument2->lpvtbl->QueryInterface(pIHTMLDocument2, @IID_IHTMLDocument3, @pIHTMLDocument3)
   pIHTMLDocument2->lpvtbl->Release(pIHTMLDocument2)
   IF hr <> S_OK OR pIHTMLDocument3 = NULL THEN RETURN ""
   ' // Get a reference to the input element
   DIM pIHTMLElement AS IHTMLElement PTR
   DIM bstrId AS AFX_BSTR = SysAllocString(**cwsId)
   hr = pIHTMLDocument3->lpvtbl->getElementById(pIHTMLDocument3, bstrId, @pIHTMLElement)
   pIHTMLDocument3->lpvtbl->Release(pIHTMLDocument3)
   SysFreeString bstrId
   IF hr <> S_OK OR pIHTMLElement = NULL THEN RETURN ""
   ' // Get the inner html text
   DIM bstrHtml AS AFX_BSTR
   hr = pIHTMLElement->lpvtbl->get_innerHtml(pIHTMLElement, @bstrHtml)
   pIHTMLElement->lpvtbl->Release(pIHTMLElement)
   DIM cwsHtml AS CWSTR = *bstrHtml
   SysFreeString bstrHtml
   RETURN cwsHtml
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the HTML between the start and end tags of the object.
' Parameters:
' - cbsId = The identifier.
' - cbsHtml = The html text to set.
' Remarks:
'   This method performs a case insensitive property search.
'   If two or more attributes have the same name (differing only in uppercase and lowercase
'   letters) this function sets values only for the last attribute created with this
'   name, and ignores all other attributes with the same name.
' Return Value:
' - S_OK if successful, or an error value otherwise.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.SetElementInnerHtmlById (BYREF cwsId AS CWSTR, BYREF cwsHtml AS CWSTR) AS HRESULT
   CWBX_DP("CWebCtx SetElementInnerHtmlById")
   IF m_pWebBrowser = NULL THEN RETURN E_POINTER
   ' // Get a reference to the IHTMLDocument2 interface
   DIM pIHTMLDocument2 AS IHTMLDocument2 PTR
   DIM hr AS HRESULT = m_pWebBrowser->get_Document(@CAST(ANY PTR, pIHTMLDocument2))
   IF hr <> S_OK OR pIHTMLDocument2 = NULL THEN RETURN hr
   ' // Get a reference to the IHTMLDocument3 interface
   DIM pIHTMLDocument3 AS IHTMLDocument3 PTR
   hr = pIHTMLDocument2->lpvtbl->QueryInterface(pIHTMLDocument2, @IID_IHTMLDocument3, @pIHTMLDocument3)
   pIHTMLDocument2->lpvtbl->Release(pIHTMLDocument2)
   IF hr <> S_OK OR pIHTMLDocument3 = NULL THEN RETURN hr
   ' // Get a reference to the input element
   DIM pIHTMLElement AS IHTMLElement PTR
   DIM bstrId AS AFX_BSTR = SysAllocString(**cwsId)
   hr = pIHTMLDocument3->lpvtbl->getElementById(pIHTMLDocument3, bstrId, @pIHTMLElement)
   pIHTMLDocument3->lpvtbl->Release(pIHTMLDocument3)
   SysFreeString bstrId
   IF hr <> S_OK OR pIHTMLElement = NULL THEN RETURN hr
   ' // Set the inner html text
   DIM bstrHtml AS AFX_BSTR = SysAllocString(**cwsHtml)
   hr = pIHTMLElement->lpvtbl->put_innerHtml(pIHTMLElement, bstrHtml)
   pIHTMLElement->lpvtbl->Release(pIHTMLElement)
   SysFreeString bstrHtml
   RETURN hr
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the focus in the specified element.
' Parameter:
' - cbsId = The identifier.
' Return value:
' - S_OK if successful, or an error value otherwise.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.SetElementFocusById (BYREF cwsId AS CWSTR) AS HRESULT
   CWBX_DP("CWebCtx SetElementFocusById")
   IF m_pWebBrowser = NULL THEN RETURN E_POINTER
   ' // Get a reference to the IHTMLDocument2 interface
   DIM pIHTMLDocument2 AS IHTMLDocument2 PTR
   DIM hr AS HRESULT = m_pWebBrowser->get_Document(@CAST(ANY PTR, pIHTMLDocument2))
   IF hr <> S_OK OR pIHTMLDocument2 = NULL THEN RETURN hr
   ' // Get a reference to the IHTMLDocument3 interface
   DIM pIHTMLDocument3 AS IHTMLDocument3 PTR
   hr = pIHTMLDocument2->lpvtbl->QueryInterface(pIHTMLDocument2, @IID_IHTMLDocument3, @pIHTMLDocument3)
   pIHTMLDocument2->lpvtbl->Release(pIHTMLDocument2)
   IF hr <> S_OK OR pIHTMLDocument3 = NULL THEN RETURN hr
   ' // Get a reference to the input element
   DIM pIHTMLElement AS IHTMLElement PTR
   DIM bstrId AS AFX_BSTR = SysAllocString(**cwsId)
   hr = pIHTMLDocument3->lpvtbl->getElementById(pIHTMLDocument3, bstrId, @pIHTMLElement)
   pIHTMLDocument3->lpvtbl->Release(pIHTMLDocument3)
   SysFreeString bstrId
   IF hr <> S_OK OR pIHTMLElement = NULL THEN RETURN hr
   ' // Get a reference to the IHTMLElement2 interface
   DIM pIHTMLElement2 AS IHTMLElement2 PTR
   hr = pIHTMLElement->lpvtbl->QueryInterface(pIHTMLElement, @IID_IHTMLElement2, @pIHTMLElement2)
   pIHTMLElement->lpvtbl->Release(pIHTMLElement)
   IF hr = S_OK AND pIHTMLElement2 <> NULL THEN
      ' // Set the focus in the element
      hr = pIHTMLElement2->lpvtbl->focus(pIHTMLElement2)
      pIHTMLElement2->lpvtbl->Release(pIHTMLElement2)
   END IF
   RETURN hr
END FUNCTION
' ========================================================================================

' ========================================================================================
'                  *** IWebBrowser2 interface additional methods ***
' ========================================================================================

' ========================================================================================
' Sets or gets a value that indicates whether the object can display dialog boxes.
' VARIANT_FALSE Dialog boxes and messages can be displayed.
' VARIANT_TRUE  Dialog boxes are not displayed.
' Note: Critical errors and security alerts are not suppressed.
' See Community additions:https://msdn.microsoft.com/en-us/library/aa768269(v=vs.85).aspx
' Effects of this property on SSL client authentication
' Setting this property to VARIANT_TRUE on the WebBrowser control has an unexpected side
' effect: SSL client authentication no longer works. In the case where you have multiple
' suitable client certificates present in the certificate store this is obvious: the
' WebBrowser control needs to display a dialog box (probably by calling InternetErrorDlg)
' to ask the user which certificate to send to the web server but the application indicated
' it doesn't want the dialog box to be displayed by setting Silent to VARIANT_TRUE.
' However, if IE's security setting "Don't prompt for client certificate selection when no
' certificates or only one certificate exists" is enabled and there is only one suitable
' client certificate present in the certificate store, the client certificate still isn't
' sent to the web server. And that, in my opinion, is a bug. If the Silent property is set
' to VARIANT_FALSE then the client certificate is sent to the web server without any
' dialog boxes being displayed - just as expected.
' ========================================================================================
PRIVATE PROPERTY CWebCtx.Silent (BYVAL bSilent AS VARIANT_BOOL)
   CWBX_DP("CWebCtx put_Silent")
   IF m_pWebBrowser THEN m_pWebBrowser->put_Silent(bSilent)
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CWebCtx.Silent () AS VARIANT_BOOL
   CWBX_DP("CWebCtx get_Silent")
   DIM bSilent AS VARIANT_BOOL
   IF m_pWebBrowser THEN m_pWebBrowser->get_Silent(@bSilent)
   RETURN bSilent
END PROPERTY
' ========================================================================================

' ========================================================================================
' Sets or gets a value that indicates whether the object can be resized.
' This property seems to be ignored by the embedded WebBrowser control.
' ========================================================================================
'PRIVATE PROPERTY CWebCtx.Resizable (BYVAL bResize AS VARIANT_BOOL)
'   CWBX_DP("CWebCtx put_Resizable")
'   IF m_pWebBrowser THEN m_pWebBrowser->put_Resizable(bResize)
'END PROPERTY
'' ========================================================================================
'' ========================================================================================
'PRIVATE PROPERTY CWebCtx.Resizable () AS VARIANT_BOOL
'   CWBX_DP("CWebCtx get_Resizable")
'   DIM bResize AS VARIANT_BOOL
'   IF m_pWebBrowser THEN m_pWebBrowser->get_Resizable(@bResize)
'   RETURN bResize
'END PROPERTY
'' ========================================================================================

' ========================================================================================
' Navigates backward one item in the history list.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.GoBack () AS HRESULT
   CWBX_DP("CWebCtx GoBack")
   IF m_pWebBrowser THEN RETURN m_pWebBrowser->GoBack
END FUNCTION
' ========================================================================================
' ========================================================================================
' Navigates forward one item in the history list.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.GoForward () AS HRESULT
   CWBX_DP("CWebCtx GoForward")
   IF m_pWebBrowser THEN RETURN m_pWebBrowser->GoForward
END FUNCTION
' ========================================================================================
' ========================================================================================
' Navigates to the current home or start page.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.GoHome () AS HRESULT
   CWBX_DP("CWebCtx GoHome")
   IF m_pWebBrowser THEN RETURN m_pWebBrowser->GoHome
END FUNCTION
' ========================================================================================
' ========================================================================================
' Navigates to the default search page of the current user.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.GoSearch () AS HRESULT
   CWBX_DP("CWebCtx GoSearch")
   IF m_pWebBrowser THEN RETURN m_pWebBrowser->GoSearch
END FUNCTION
' ========================================================================================
' ========================================================================================
' Reloads the file that is currently displayed in the object.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.Refresh () AS HRESULT
   CWBX_DP("CWebCtx Refresh")
   IF m_pWebBrowser THEN RETURN m_pWebBrowser->Refresh
END FUNCTION
' ========================================================================================
' ========================================================================================
' Reloads the file that is currently displayed with the specified refresh level.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.Refresh2 (BYVAL nLevel AS RefreshConstants) AS HRESULT
   CWBX_DP("CWebCtx Refresh2")
   DIM vLevel AS VARIANT
   vLevel.vt = VT_I4 : vLevel.lVal = nLevel
   IF m_pWebBrowser THEN RETURN m_pWebBrowser->Refresh2(@vLevel)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Cancels a pending navigation or download, and stops dynamic page elements, such as
' background sounds and animations.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.Stop () AS HRESULT
   CWBX_DP("CWebCtx Stop")
   IF m_pWebBrowser THEN RETURN m_pWebBrowser->Stop
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets the automation object of the active document, if any.
' ========================================================================================
PRIVATE PROPERTY CWebCtx.Document () AS IHtmlDocument2 PTR
   CWBX_DP("CWebCtx Document")
   DIM pdisp AS IDispatch PTR
   IF m_pWebBrowser THEN m_pWebBrowser->get_Document(cast(ANY PTR, @pdisp))
   PROPERTY = cast(IHtmlDocument2 PTR, pdisp)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Retrieves the path or title of the resource that is currently displayed.
' ========================================================================================
PRIVATE PROPERTY CWebCtx.LocationName () AS CWSTR
   CWBX_DP("CWebCtx get_LocationName")
   DIM bstrName AS AFX_BSTR
   IF m_pWebBrowser THEN m_pWebBrowser->get_LocationName(@bstrName)
   PROPERTY = *bstrName
   SysFreestring bstrName
END PROPERTY
' ========================================================================================
' ========================================================================================
' Retrieves the URL of the resource that is currently displayed.
' ========================================================================================
PRIVATE PROPERTY CWebCtx.LocationURL () AS CWSTR
   CWBX_DP("CWebCtx get_LocationURL")
   DIM bstrName AS AFX_BSTR
   IF m_pWebBrowser THEN m_pWebBrowser->get_LocationURL(@bstrName)
   PROPERTY = *bstrName
   SysFreestring bstrName
END PROPERTY
' ========================================================================================
' ========================================================================================
' TRUE if the object is engaged in a navigation or downloading operation; FALSE, otherwise.
' ========================================================================================
PRIVATE PROPERTY CWebCtx.Busy () AS BOOLEAN
   CWBX_DP("CWebCtx Busy")
   DIM bBusy AS VARIANT_BOOL
   IF m_pWebBrowser THEN m_pWebBrowser->get_Busy(@bBusy)
   PROPERTY = bBusy
END PROPERTY
' ========================================================================================
' ========================================================================================
' Queries the object for the status of commands using the IOleCommandTarget::QueryStatus method.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.QueryStatusWB (BYVAL cmdID AS OLECMDID) AS OLECMDF
   CWBX_DP("CWebCtx QueryStatusWB")
   DIM cmdf AS OLECMDF
   IF m_pWebBrowser THEN m_pWebBrowser->QueryStatusWB(cmdID, @cmdf)
   RETURN cmdf
END FUNCTION
' ========================================================================================
' ========================================================================================
' Executes a command and returns the status of the command execution using the IOleCommandTarget interface.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.ExecWB (BYVAL cmdID AS OLECMDID, BYVAL cmdexecopt AS OLECMDEXECOPT, BYVAL pvaIn AS VARIANT PTR = NULL, BYVAL pvaOut AS VARIANT PTR = NULL) AS HRESULT
   CWBX_DP("CWebCtx ExecWB")
   IF m_pWebBrowser THEN RETURN(m_pWebBrowser->ExecWB(cmdID, cmdexecopt, pvaIn, pvaOut))
END FUNCTION
' ========================================================================================
' ========================================================================================
' ========================================================================================
' Gets the ready state of the object.
' Return value:
' READYSTATE_UNINITIALIZED = 0
'   Default initialization state.
' READYSTATE_LOADING = 1
'   Object is currently loading its properties.
' READYSTATE_LOADED = 2
'   Object has been initialized.
' READYSTATE_INTERACTIVE = 3
'   Object is interactive, but not all of its data is available.
' READYSTATE_COMPLETE = 4
'   Object has received all of its data.
' ========================================================================================
PRIVATE PROPERTY CWebCtx.ReadyState () AS tagREADYSTATE
   CWBX_DP("CWebCtx ReadyState")
   DIM nState AS tagREADYSTATE
   IF m_pWebBrowser THEN m_pWebBrowser->get_ReadyState(@nState)
   PROPERTY = nState
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets or gets a value that indicates whether the object is registered as a top-level browser window.
' ========================================================================================
PRIVATE PROPERTY CWebCtx.RegisterAsBrowser () AS BOOLEAN
   CWBX_DP("CWebCtx get_RegisterAsBrowser")
   DIM bRegister AS VARIANT_BOOL
   IF m_pWebBrowser THEN m_pWebBrowser->get_RegisterAsBrowser(@bRegister)
   PROPERTY = bRegister
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CWebCtx.RegisterAsBrowser (BYVAL bRegister AS BOOLEAN)
   CWBX_DP("CWebCtx put_RegisterAsBrowser")
   IF m_pWebBrowser THEN m_pWebBrowser->put_RegisterAsBrowser(bRegister)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Sets or gets a value that indicates whether the object is registered as a drop target for navigation.
' ========================================================================================
PRIVATE PROPERTY CWebCtx.RegisterAsDropTarget () AS BOOLEAN
   CWBX_DP("CWebCtx get_RegisterAsDropTarget")
   DIM bRegister AS VARIANT_BOOL
   IF m_pWebBrowser THEN m_pWebBrowser->get_RegisterAsDropTarget(@bRegister)
   PROPERTY = bRegister
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CWebCtx.RegisterAsDropTarget (BYVAL bRegister AS BOOLEAN)
   CWBX_DP("CWebCtx put_RegisterAsDropTarget")
   IF m_pWebBrowser THEN m_pWebBrowser->put_RegisterAsDropTarget(bRegister)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Activates the print preview dialog.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.PrintPreview () AS HRESULT
   IF m_pWebBrowser THEN RETURN  m_pWebBrowser->ExecWB(OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_PROMPTUSER, NULL, NULL)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Activates de print dialog.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.PrintPage (BYVAL bPromptUser AS BOOLEAN = TRUE) AS HRESULT
   IF m_pWebBrowser = NULL THEN RETURN E_POINTER
   IF bPromptUser THEN
      m_pWebBrowser->ExecWB(OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER, NULL, NULL)
   ELSE
      m_pWebBrowser->ExecWB(OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER, NULL, NULL)
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Activates the page setup dialog.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.PageSetUp () AS HRESULT
   IF m_pWebBrowser THEN RETURN  m_pWebBrowser->ExecWB(OLECMDID_PAGESETUP, OLECMDEXECOPT_PROMPTUSER, NULL, NULL)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Activates the properties dialog.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.PageProperties () AS HRESULT
   IF m_pWebBrowser THEN RETURN  m_pWebBrowser->ExecWB(OLECMDID_PROPERTIES, OLECMDEXECOPT_PROMPTUSER, NULL, NULL)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Activates the save file dialog.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.SaveAs () AS HRESULT
   IF m_pWebBrowser THEN RETURN  m_pWebBrowser->ExecWB(OLECMDID_SAVEAS, OLECMDEXECOPT_PROMPTUSER, NULL, NULL)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Activates the find dialog.
' Warning: This code uses an undocumented command-group GUID that is subject to change in
' the future. Currently it works in all versions of Internet Explorer up to 11.
' See http://support.microsoft.com/?kbid=311288
' ========================================================================================
PRIVATE FUNCTION CWebCtx.Find () AS HRESULT
   IF m_pWebBrowser = NULL THEN RETURN E_POINTER
   DIM pIHTMLDocument2 AS IHTMLDocument2 PTR
   DIM hr AS HRESULT = m_pWebBrowser->get_Document(@CAST(ANY PTR, pIHTMLDocument2))
   IF hr <> S_OK OR pIHTMLDocument2 = NULL THEN RETURN hr
   DIM CGID_WebBrowser_ AS GUID = (&hED016940, &hBD5B, &h11CF, {&hBA, &h4E, &h00, &hC0, &h4F, &hD7, &h08, &h16})
   DIM pCmdTarget AS IOleCommandTarget PTR
   hr = pIHTMLDocument2->lpvtbl->QueryInterface(pIHTMLDocument2, @IID_IOleCommandTarget, @pCmdTarget)
   pIHTMLDocument2->lpvtbl->Release(pIHTMLDocument2)
   IF hr <> S_OK OR pIHTMLDocument2 = NULL THEN RETURN hr
   hr = pCmdTarget->lpvtbl->Exec(pCmdTarget, @CGID_WebBrowser_, 1, 0, NULL, NULL)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Displays the source code of the page in an instance of NotePad.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.ShowSource () AS HRESULT
   IF m_pWebBrowser = NULL THEN RETURN E_POINTER
   DIM pIHTMLDocument2 AS IHTMLDocument2 PTR
   DIM hr AS HRESULT = m_pWebBrowser->get_Document(@CAST(ANY PTR, pIHTMLDocument2))
   IF hr <> S_OK OR pIHTMLDocument2 = NULL THEN RETURN hr
   DIM CGID_WebBrowser_ AS GUID = (&hED016940, &hBD5B, &h11CF, {&hBA, &h4E, &h00, &hC0, &h4F, &hD7, &h08, &h16})
   DIM pCmdTarget AS IOleCommandTarget PTR
   hr = pIHTMLDocument2->lpvtbl->QueryInterface(pIHTMLDocument2, @IID_IOleCommandTarget, @pCmdTarget)
   pIHTMLDocument2->lpvtbl->Release(pIHTMLDocument2)
   IF hr <> S_OK OR pIHTMLDocument2 = NULL THEN RETURN hr
   hr = pCmdTarget->lpvtbl->Exec(pCmdTarget, @CGID_WebBrowser_, 2, 0, NULL, NULL)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Activates the Internet options dialog.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.InternetOptions () AS HRESULT
   IF m_pWebBrowser = NULL THEN RETURN E_POINTER
   DIM pIHTMLDocument2 AS IHTMLDocument2 PTR
   DIM hr AS HRESULT = m_pWebBrowser->get_Document(@CAST(ANY PTR, pIHTMLDocument2))
   IF hr <> S_OK OR pIHTMLDocument2 = NULL THEN RETURN hr
   DIM CGID_WebBrowser_ AS GUID = (&hED016940, &hBD5B, &h11CF, {&hBA, &h4E, &h00, &hC0, &h4F, &hD7, &h08, &h16})
   DIM pCmdTarget AS IOleCommandTarget PTR
   hr = pIHTMLDocument2->lpvtbl->QueryInterface(pIHTMLDocument2, @IID_IOleCommandTarget, @pCmdTarget)
   pIHTMLDocument2->lpvtbl->Release(pIHTMLDocument2)
   IF hr <> S_OK OR pIHTMLDocument2 = NULL THEN RETURN hr
   hr = pCmdTarget->lpvtbl->Exec(pCmdTarget, @CGID_WebBrowser_, 3, 0, NULL, NULL)
END FUNCTION
' ========================================================================================


' ########################################################################################
'                              *** HELPER PROCEDURES ***
' ########################################################################################

' ========================================================================================
' Returns a pointer to the hosted WebBrowser control given the handle of the form, or any
' control in the form, and the control identifier.
' Parameters:
' - hwnd = Reference to a window handle.
' - Control identifier, e.g. IDC_WEBBROWSER.
' Returns a pointer to the IWeBbrowser2 interface or NULL.
' ========================================================================================
PRIVATE FUNCTION AfxGetBrowserPtr (BYVAL hwnd AS HWND, BYVAL cID AS WORD) AS Afx_IWebBrowser2 PTR
   DIM pAxHost AS CAxHost PTR = AfxCAxHostPtr(hwnd, cID)
   IF pAxHost THEN RETURN CAST(Afx_IWebBrowser2 PTR, cast(ULONG_PTR, pAxHost->m_pOcx))
   RETURN NULL
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes one or more HTML expressions to a document.
' Parameters
' - cbsHtml: Text and HTML tags to write.
' - cr: Write the HTML text followed by a carriage return.
' Remarks
'   In HTML, the carriage return is ignored unless it occurs in preformatted text.
'   Note  When document.IHTMLDocument2::write or document.IHTMLDocument2::writeln is used
'   in an event handler, you must also use document.IHTMLDocument2::close.
' Return value:
' - S_OK if successful, or an error value otherwise.
' ========================================================================================
PRIVATE FUNCTION AfxWriteHtml (BYVAL pWebBrowser AS Afx_IWebBrowser2 PTR, BYREF cwsHtml AS CWSTR, BYVAL cr AS BOOLEAN = FALSE) AS HRESULT

   IF pWebBrowser = NULL THEN RETURN E_POINTER
   DIM pIHTMLDocument2 AS IHTMLDocument2 PTR
   DIM hr AS HRESULT = pWebBrowser->get_Document(@cast(ANY PTR, pIHTMLDocument2))
   IF hr <> S_OK OR pIHTMLDocument2 = NULL THEN RETURN hr

   ' // Create a safearray of variants of one element
   DIM psarray AS SafeArray PTR
   DIM rgsabounds(0) AS SAFEARRAYBOUND = {(1, 0)}
   psarray = SafeArrayCreate(VT_VARIANT, 1, @rgsabounds(0))
   IF psarray = NULL THEN RETURN E_FAIL

   ' // Fill the safearray with the script
   DIM ix AS LONG = 0
   DIM vHtml AS VARIANT
   vHtml.vt = VT_BSTR
   vHtml.bstrVal = SysAllocString(**cwsHtml)
   SafeArrayPutElement(psarray, @ix, @vHtml)
   ' // Write the string
   IF cr THEN
      hr = pIHTMLDocument2->lpvtbl->writeln(pIHTMLDocument2, psarray)
   ELSE
      hr = pIHTMLDocument2->lpvtbl->write(pIHTMLDocument2, psarray)
   END IF
   ' // Destroy the safearray
   SafeArrayDestroy psarray
   ' // Free the pIHTMLDocument2 interface
   pIHTMLDocument2->lpvtbl->Release(pIHTMLDocument2)
   ' // Clear the variant
   VariantClear @vHtml
   RETURN hr

END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the ID of the active element (the object that has the focus when the parent
' document has focus).
' Return value: The ID of the active element.
' ========================================================================================
PRIVATE FUNCTION AfxGetActiveElementId (BYVAL pWebBrowser AS Afx_IWebBrowser2 PTR) AS CWSTR
   IF pWebBrowser = NULL THEN RETURN ""
   DIM pIHTMLDocument2 AS IHTMLDocument2 PTR
   DIM hr AS HRESULT = pWebBrowser->get_Document(@cast(ANY PTR, pIHTMLDocument2))
   IF hr <> S_OK OR pIHTMLDocument2 = NULL THEN RETURN ""
   DIM pIHTMLElement AS IHTMLElement PTR
   hr = pIHTMLDocument2->lpvtbl->get_activeElement(pIHTMLDocument2, @pIHTMLElement)
   pIHTMLDocument2->lpvtbl->Release(pIHTMLDocument2)
   IF hr <> S_OK OR pIHTMLElement = NULL THEN RETURN ""
   DIM bstrId AS AFX_BSTR
   hr = pIHTMLElement->lpvtbl->get_id(pIHTMLElement, @bstrId)
   pIHTMLElement->lpvtbl->Release(pIHTMLElement)
   DIM cwsId AS CWSTR = *bstrId
   SysFreeString bstrId
   RETURN cwsId
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the value attribute of the specified identifier.
' Parameter:
' - cbsId = The identifier.
' Return Value:
' - A variant containing the value as defined by the attribute.
' Remarks:
'   This method performs a case insensitive property search.
'   If two or more attributes have the same name (differing only in uppercase and lowercase
'   letters) this function retrieves values only for the last attribute created with this
'   name, and ignores all other attributes with the same name.
' ========================================================================================
PRIVATE FUNCTION AfxGetElementValueById (BYVAL pWebBrowser AS Afx_IWebBrowser2 PTR, BYREF cwsId AS CWSTR) AS VARIANT
   IF pWebBrowser = NULL THEN EXIT FUNCTION
   ' // Get a reference to the IHTMLDocument2 interface
   DIM pIHTMLDocument2 AS IHTMLDocument2 PTR
   DIM hr AS HRESULT = pWebBrowser->get_Document(@cast(ANY PTR, pIHTMLDocument2))
   IF hr <> S_OK OR pIHTMLDocument2 = NULL THEN EXIT FUNCTION
   ' // Get a reference to the IHTMLDocument3 interface
   DIM pIHTMLDocument3 AS IHTMLDocument3 PTR
   hr = pIHTMLDocument2->lpvtbl->QueryInterface(pIHTMLDocument2, @IID_IHTMLDocument3, @pIHTMLDocument3)
   pIHTMLDocument2->lpvtbl->Release(pIHTMLDocument2)
   IF hr <> S_OK OR pIHTMLDocument3 = NULL THEN EXIT FUNCTION
   ' // Get a reference to the input element
   DIM pIHTMLElement AS IHTMLElement PTR
   DIM bstrId AS AFX_BSTR = SysAllocString(**cwsId)
   hr = pIHTMLDocument3->lpvtbl->getElementById(pIHTMLDocument3, bstrId, @pIHTMLElement)
   pIHTMLDocument3->lpvtbl->Release(pIHTMLDocument3)
   SysFreeString bstrId
   IF hr <> S_OK OR pIHTMLElement = NULL THEN EXIT FUNCTION
   ' // Get the value
   DIM vValue AS VARIANT, bstrAttr AS AFX_BSTR = SysAllocString("value")
   hr = pIHTMLElement->lpvtbl->getAttribute(pIHTMLElement, bstrAttr, 0, @vValue)
   pIHTMLElement->lpvtbl->Release(pIHTMLElement)
   SysFreeString bstrAttr
   RETURN vValue
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the value attribute of the specified identifier.
' Parameter:
' - cbsId = The identifier.
' - cvValue = Variant that specifies the string, number, or Boolean to assign to the attribute.
' Remarks:
'   This method performs a case insensitive property search.
'   If two or more attributes have the same name (differing only in uppercase and lowercase
'   letters) this function sets values only for the last attribute created with this
'   name, and ignores all other attributes with the same name.
' Return value:
' - S_OK if successful, or an error value otherwise.
' ========================================================================================
PRIVATE FUNCTION AfxSetElementValueById (BYVAL pWebBrowser AS Afx_IWebBrowser2 PTR, BYREF cwsId AS CWSTR, BYVAL vValue AS VARIANT) AS HRESULT
   IF pWebBrowser = NULL THEN RETURN E_POINTER
   ' // Get a reference to the IHTMLDocument2 interface
   DIM pIHTMLDocument2 AS IHTMLDocument2 PTR
   DIM hr AS HRESULT = pWebBrowser->get_Document(@cast(ANY PTR, pIHTMLDocument2))
   IF hr <> S_OK OR pIHTMLDocument2 = NULL THEN RETURN hr
   ' // Get a reference to the IHTMLDocument3 interface
   DIM pIHTMLDocument3 AS IHTMLDocument3 PTR
   hr = pIHTMLDocument2->lpvtbl->QueryInterface(pIHTMLDocument2, @IID_IHTMLDocument3, @pIHTMLDocument3)
   pIHTMLDocument2->lpvtbl->Release(pIHTMLDocument2)
   IF hr <> S_OK OR pIHTMLDocument3 = NULL THEN RETURN hr
   ' // Get a reference to the input element
   DIM pIHTMLElement AS IHTMLElement PTR
   DIM bstrId AS AFX_BSTR = SysAllocString(**cwsId)
   hr = pIHTMLDocument3->lpvtbl->getElementById(pIHTMLDocument3, bstrId, @pIHTMLElement)
   pIHTMLDocument3->lpvtbl->Release(pIHTMLDocument3)
   SysFreeString bstrId
   IF hr <> S_OK OR pIHTMLElement = NULL THEN RETURN hr
   ' // Set the value
   DIM bstrAttr AS AFX_BSTR = SysAllocString("value")
   hr = pIHTMLElement->lpvtbl->setAttribute(pIHTMLElement, bstrAttr, vValue, 0)
   pIHTMLElement->lpvtbl->Release(pIHTMLElement)
   SysFreestring bstrAttr
   RETURN hr
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the HTML between the start and end tags of the object.
' Parameter:
' - cbsId = The idenfier.
' Return value:
' - A string containing the HTML text.
' Remarks:
'   This method performs a case insensitive property search.
'   If two or more attributes have the same name (differing only in uppercase and lowercase
'   letters) this function retrieves values only for the last attribute created with this
'   name, and ignores all other attributes with the same name.
' ========================================================================================
PRIVATE FUNCTION AfxGetElementInnerHtmlById (BYVAL pWebBrowser AS Afx_IWebBrowser2 PTR, BYREF cwsId AS CWSTR) AS CWSTR
   IF pWebBrowser = NULL THEN RETURN ""
   ' // Get a reference to the IHTMLDocument2 interface
   DIM pIHTMLDocument2 AS IHTMLDocument2 PTR
   DIM hr AS HRESULT = pWebBrowser->get_Document(@cast(ANY PTR, pIHTMLDocument2))
   IF hr <> S_OK OR pIHTMLDocument2 = NULL THEN RETURN ""
   ' // Get a reference to the IHTMLDocument3 interface
   DIM pIHTMLDocument3 AS IHTMLDocument3 PTR
   hr = pIHTMLDocument2->lpvtbl->QueryInterface(pIHTMLDocument2, @IID_IHTMLDocument3, @pIHTMLDocument3)
   pIHTMLDocument2->lpvtbl->Release(pIHTMLDocument2)
   IF hr <> S_OK OR pIHTMLDocument3 = NULL THEN RETURN ""
   ' // Get a reference to the input element
   DIM pIHTMLElement AS IHTMLElement PTR
   DIM bstrId AS AFX_BSTR = SysAllocString(**cwsId)
   hr = pIHTMLDocument3->lpvtbl->getElementById(pIHTMLDocument3, bstrId, @pIHTMLElement)
   pIHTMLDocument3->lpvtbl->Release(pIHTMLDocument3)
   SysFreeString bstrId
   IF hr <> S_OK OR pIHTMLElement = NULL THEN RETURN ""
   ' // Get the inner html text
   DIM bstrHtml AS AFX_BSTR
   hr = pIHTMLElement->lpvtbl->get_innerHtml(pIHTMLElement, @bstrHtml)
   pIHTMLElement->lpvtbl->Release(pIHTMLElement)
   DIM cwsHtml AS CWSTR = *bstrHtml
   SysFreeString bstrHtml
   RETURN cwsHtml
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the HTML between the start and end tags of the object.
' Parameters:
' - cbsId = The identifier.
' - cbsHtml = The html text to set.
' Remarks:
'   This method performs a case insensitive property search.
'   If two or more attributes have the same name (differing only in uppercase and lowercase
'   letters) this function sets values only for the last attribute created with this
'   name, and ignores all other attributes with the same name.
' Return Value:
' - S_OK if successful, or an error value otherwise.
' ========================================================================================
PRIVATE FUNCTION AfxSetElementInnerHtmlById (BYVAL pWebBrowser AS Afx_IWebBrowser2 PTR, BYREF cwsId AS CWSTR, BYREF cwsHtml AS CWSTR) AS HRESULT
   IF pWebBrowser = NULL THEN RETURN E_POINTER
   ' // Get a reference to the IHTMLDocument2 interface
   DIM pIHTMLDocument2 AS IHTMLDocument2 PTR
   DIM hr AS HRESULT = pWebBrowser->get_Document(@cast(ANY PTR, pIHTMLDocument2))
   IF hr <> S_OK OR pIHTMLDocument2 = NULL THEN RETURN hr
   ' // Get a reference to the IHTMLDocument3 interface
   DIM pIHTMLDocument3 AS IHTMLDocument3 PTR
   hr = pIHTMLDocument2->lpvtbl->QueryInterface(pIHTMLDocument2, @IID_IHTMLDocument3, @pIHTMLDocument3)
   pIHTMLDocument2->lpvtbl->Release(pIHTMLDocument2)
   IF hr <> S_OK OR pIHTMLDocument3 = NULL THEN RETURN hr
   ' // Get a reference to the input element
   DIM pIHTMLElement AS IHTMLElement PTR
   DIM bstrId AS AFX_BSTR = SysAllocString(**cwsId)
   hr = pIHTMLDocument3->lpvtbl->getElementById(pIHTMLDocument3, bstrId, @pIHTMLElement)
   pIHTMLDocument3->lpvtbl->Release(pIHTMLDocument3)
   SysFreeString bstrId
   IF hr <> S_OK OR pIHTMLElement = NULL THEN RETURN hr
   ' // Set the inner html text
   DIM bstrHtml AS AFX_BSTR = SysAllocString(**cwsHtml)
   hr = pIHTMLElement->lpvtbl->put_innerHtml(pIHTMLElement, bstrHtml)
   pIHTMLElement->lpvtbl->Release(pIHTMLElement)
   SysFreeString bstrHtml
   RETURN hr
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the focus in the specified element.
' Parameter:
' - cbsId = The identifier.
' Return value:
' - S_OK if successful, or an error value otherwise.
' ========================================================================================
PRIVATE FUNCTION AfxSetElementFocusById (BYVAL pWebBrowser AS Afx_IWebBrowser2 PTR, BYREF cwsId AS CWSTR) AS HRESULT
   IF pWebBrowser = NULL THEN RETURN E_POINTER
   ' // Get a reference to the IHTMLDocument2 interface
   DIM pIHTMLDocument2 AS IHTMLDocument2 PTR
   DIM hr AS HRESULT = pWebBrowser->get_Document(@cast(ANY PTR, pIHTMLDocument2))
   IF hr <> S_OK OR pIHTMLDocument2 = NULL THEN RETURN hr
   ' // Get a reference to the IHTMLDocument3 interface
   DIM pIHTMLDocument3 AS IHTMLDocument3 PTR
   hr = pIHTMLDocument2->lpvtbl->QueryInterface(pIHTMLDocument2, @IID_IHTMLDocument3, @pIHTMLDocument3)
   pIHTMLDocument2->lpvtbl->Release(pIHTMLDocument2)
   IF hr <> S_OK OR pIHTMLDocument3 = NULL THEN RETURN hr
   ' // Get a reference to the input element
   DIM pIHTMLElement AS IHTMLElement PTR
   DIM bstrId AS AFX_BSTR = SysAllocString(**cwsId)
   hr = pIHTMLDocument3->lpvtbl->getElementById(pIHTMLDocument3, bstrId, @pIHTMLElement)
   pIHTMLDocument3->lpvtbl->Release(pIHTMLDocument3)
   SysFreeString bstrId
   IF hr <> S_OK OR pIHTMLElement = NULL THEN RETURN hr
   ' // Get a reference to the IHTMLElement2 interface
   DIM pIHTMLElement2 AS IHTMLElement2 PTR
   hr = pIHTMLElement->lpvtbl->QueryInterface(pIHTMLElement, @IID_IHTMLElement2, @pIHTMLElement2)
   pIHTMLElement->lpvtbl->Release(pIHTMLElement)
   IF hr = S_OK AND pIHTMLElement2 <> NULL THEN
      ' // Set the focus in the element
      hr = pIHTMLElement2->lpvtbl->focus(pIHTMLElement2)
      pIHTMLElement2->lpvtbl->Release(pIHTMLElement2)
   END IF
   RETURN hr
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a string that represents the text between the start and end body tags without
' any associated HTML.
' ========================================================================================
PRIVATE FUNCTION AfxGetBodyInnerText (BYVAL pWebBrowser AS Afx_IWebBrowser2 PTR) AS CWSTR
   IF pWebBrowser = NULL THEN RETURN ""
   DIM pIHTMLDocument2 AS IHTMLDocument2 PTR
   DIM hr AS HRESULT = pWebBrowser->get_Document(@CAST(ANY PTR, pIHTMLDocument2))
   IF hr <> S_OK OR pIHTMLDocument2 = NULL THEN RETURN ""
   DIM pIHTMLElement AS IHTMLElement PTR
   hr = pIHTMLDocument2->lpvtbl->get_body(pIHTMLDocument2, @pIHTMLElement)
   pIHTMLDocument2->lpvtbl->Release(pIHTMLDocument2)
   IF hr <> S_OK OR pIHTMLElement = NULL THEN RETURN ""
   DIM bstrText AS AFX_BSTR
   hr = pIHTMLElement->lpvtbl->get_innerText(pIHTMLElement, @bstrText)
   pIHTMLElement->lpvtbl->Release(pIHTMLElement)
   DIM cwsText AS CWSTR = *bstrText
   SysFreeString bstrText
   RETURN cwsText
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a string that represents the text and html elements between the start and end body tags.
' ========================================================================================
PRIVATE FUNCTION AfxGetBodyInnerHtml (BYVAL pWebBrowser AS Afx_IWebBrowser2 PTR) AS CWSTR
   IF pWebBrowser = NULL THEN RETURN ""
   DIM pIHTMLDocument2 AS IHTMLDocument2 PTR
   DIM hr AS HRESULT = pWebBrowser->get_Document(@CAST(ANY PTR, pIHTMLDocument2))
   IF hr <> S_OK OR pIHTMLDocument2 = NULL THEN RETURN ""
   DIM pIHTMLElement AS IHTMLElement PTR
   hr = pIHTMLDocument2->lpvtbl->get_body(pIHTMLDocument2, @pIHTMLElement)
   pIHTMLDocument2->lpvtbl->Release(pIHTMLDocument2)
   IF hr <> S_OK OR pIHTMLElement = NULL THEN RETURN ""
   DIM bstrText AS AFX_BSTR
   hr = pIHTMLElement->lpvtbl->get_innerHtml(pIHTMLElement, @bstrText)
   pIHTMLElement->lpvtbl->Release(pIHTMLElement)
   DIM cwsText AS CWSTR = *bstrText
   SysFreeString bstrText
   RETURN cwsText
END FUNCTION
' ========================================================================================

END NAMESPACE
