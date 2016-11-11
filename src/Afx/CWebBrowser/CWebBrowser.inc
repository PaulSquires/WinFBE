' ########################################################################################
' Microsoft Windows
' File: CWebBrowser.inc
' Contents: WebBrowser control wrapper class
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "Afx/AfxExDisp.bi"
#include once "Afx/COleCon/COleCon.inc"
#include once "Afx/AfxCOM.inc"

NAMESPACE Afx

' ========================================================================================
' Macro for debug
' To allow debugging, define _CWB_DEBUG_ 1 in your application before including this file.
' ========================================================================================
#ifndef _CWB_DEBUG_
   #define _CWB_DEBUG_ 0
#endif
#ifndef _CWB_DP_
   #define _CWB_DP_ 1
   #MACRO CWB_DP(st)
      #if (_CWB_DEBUG_ = 1)
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

TYPE CWebBrowserEventsImpl_ AS CWebBrowserEventsImpl

' // DocHostUIHandler events class
#include once "Afx/CWebBrowser/CDocHostUIHandlerImpl.inc"
' // HtmlDocumentEvents2 class
#include once "Afx/CWebBrowser/CHtmlDocumentEventsImpl.inc"
' // WebBrowser events class
#include once "Afx/CWebBrowser/CWebBrowserEventsImpl.inc"

' ========================================================================================
' CWebBrowser class
' ========================================================================================
TYPE CWebBrowser

   m_pOleCon AS COleCon PTR                                 ' // Pointer to the OLE container
   m_pWebBrowser AS Afx_IWebBrowser2 PTR                    ' // Pointer to the IWebBrowser interface
   m_pEvtObj AS CWebBrowserEventsImpl PTR                   ' // Pointer to the implemented Web Browser events class
   m_pDocHostUIHandler AS CDocHostUIHandler2Impl PTR        ' // Pointer to the implemented DocHostUIHandler interface

   DECLARE CONSTRUCTOR (BYVAL pWindow AS CWindow PTR, BYVAL cID AS INTEGER, _
      BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
      BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, BYVAL ambientFlags AS DWORD = 0)
   DECLARE DESTRUCTOR
   DECLARE FUNCTION OleConPtr () AS COleCon PTR
   DECLARE FUNCTION BrowserPtr () AS Afx_IWebBrowser2 PTR
   DECLARE FUNCTION hWindow () AS HWND
   DECLARE FUNCTION Navigate (BYVAL pwszUrl AS WSTRING PTR, BYVAL Flags AS VARIANT PTR = NULL, BYVAL TargetFrameName AS VARIANT PTR = NULL, BYVAL PostData AS VARIANT PTR = NULL, BYVAL Headers AS VARIANT PTR = NULL) AS HRESULT
   DECLARE FUNCTION Navigate (BYVAL vUrl AS VARIANT PTR, BYVAL Flags AS VARIANT PTR = NULL, BYVAL TargetFrameName AS VARIANT PTR = NULL, BYVAL PostData AS VARIANT PTR = NULL, BYVAL Headers AS VARIANT PTR = NULL) AS HRESULT
   DECLARE FUNCTION Advise () AS HRESULT
   DECLARE FUNCTION Unadvise () AS HRESULT
   DECLARE FUNCTION SetUIHandler () AS HRESULT
   DECLARE FUNCTION SetEventProc (BYVAL pwszEventName AS WSTRING PTR, BYVAL pProc AS ANY PTR) AS HRESULT
   DECLARE FUNCTION SetUIEventProc (BYVAL pwszEventName AS WSTRING PTR, BYVAL pProc AS ANY PTR) AS HRESULT

   ' // IWebBrowser2 interface
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

   ' // IHtmlDoc2 wrappers
   DECLARE FUNCTION WriteHtml (BYREF cwsHtml AS CWSTR, BYVAL cr AS BOOLEAN = FALSE) AS HRESULT
   DECLARE FUNCTION GetActiveElementId () AS CWSTR
   DECLARE FUNCTION GetElementValueById (BYREF cwsId AS CWSTR) AS VARIANT
   DECLARE FUNCTION SetElementValueById (BYREF cwsId AS CWSTR, BYVAL vValue AS VARIANT) AS HRESULT
   DECLARE FUNCTION GetElementInnerHtmlById (BYREF cwsId AS CWSTR) AS CWSTR
   DECLARE FUNCTION SetElementInnerHtmlById (BYREF cwsId AS CWSTR, BYREF cwsHtml AS CWSTR) AS HRESULT
   DECLARE FUNCTION SetElementFocusById (BYREF cwsId AS CWSTR) AS HRESULT

END TYPE
' ========================================================================================

' ========================================================================================
' Constructor
' ========================================================================================
CONSTRUCTOR CWebBrowser (BYVAL pWindow AS CWindow PTR, BYVAL cID AS INTEGER, _
BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, BYVAL ambientFlags AS DWORD = 0)
   CWB_DP("CONSTRUCTOR CWebBrowser")

   ' // Create an instance of the OLE container
   m_pOleCon = NEW COleCon(pWindow, cID, "Shell.Explorer", x, y, nWidth, nHeight, dwStyle, dwExStyle, ambientFlags)
   ' // Get a reference to the IWebBrowser2 interface
   IF m_pOleCon THEN m_pWebBrowser = cast(Afx_IWebBrowser2 PTR, cast(ULONG_PTR, m_pOleCon->OcxDispPtr))
   ' // Navigate to a blank page to allow customization by calling the SetUIHandler method.
   ' // Otherwise, it will fail when trying to get the IHtmlDocument2 interface.
   IF m_pWebBrowser THEN
      DIM vUrl AS VARIANT : vUrl.vt = VT_BSTR : vUrl.bstrVal = SysAllocString("about:blank")
      m_pWebBrowser->Navigate2(@vUrl, NULL, NULL, NULL, NULL)
      VariantClear @vurl
   END IF

END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destructor
' ========================================================================================
DESTRUCTOR CWebBrowser
   CWB_DP("BEGIN DESTRUCTOR CWebBrowser")
   ' // Disconnect events and delete the OLE container
   IF m_pOleCon THEN
      m_pOleCon->Unadvise
      Delete m_pOleCon
   END IF
   ' // Delete the WebBrowser event's class
   ' // Important: Don't delete it before deleting the OLE container because
   ' // it deletes the IDocHostUIHandler class and you will get a GPF when the
   ' // OLE container will call IDocHostUIHandler to send a message.
   IF m_pEvtObj THEN Delete m_pEvtObj
   IF m_pDocHostUIHandler THEN Delete m_pDocHostUIHandler
   CWB_DP("END DESTRUCTOR CWebBrowser")
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns a pointer to the OLE container class.
' ========================================================================================
PRIVATE FUNCTION CWebBrowser.OleConPtr () AS COleCon PTR
   CWB_DP("CWebBrowser.OleConPtr")
   RETURN m_pOleCon
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a direct pointer to the IWebBrowser2 interface of the hosted WebBrowser control.
' ========================================================================================
PRIVATE FUNCTION CWebBrowser.BrowserPtr () AS Afx_IWebBrowser2 PTR
   CWB_DP("CWebBrowser.BrowserPtr")
   RETURN m_pWebBrowser
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the handle of the OLE container hosting window
' ========================================================================================
PRIVATE FUNCTION CWebBrowser.hWindow () AS HWND
   CWB_DP("CWebBrowser.hWindow")
   RETURN m_pOleCon->hWindow
END FUNCTION
' ========================================================================================

' ========================================================================================
' Navigates to the specified URL.
' ========================================================================================
PRIVATE FUNCTION CWebBrowser.Navigate (BYVAL pwszUrl AS WSTRING PTR, BYVAL Flags AS VARIANT PTR = NULL, BYVAL TargetFrameName AS VARIANT PTR = NULL, BYVAL PostData AS VARIANT PTR = NULL, BYVAL Headers AS VARIANT PTR = NULL) AS HRESULT
   CWB_DP("CWebBrowser.Navigate - pwszUrl")
   DIM vUrl AS VARIANT : vUrl.vt = VT_BSTR : vUrl.bstrVal = SysAllocString(pwszUrl)
   IF m_pWebBrowser THEN FUNCTION = m_pWebBrowser->Navigate2(@vUrl, Flags, TargetFrameName, PostData, Headers)
   VariantClear(@vUrl)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CWebBrowser.Navigate (BYVAL vUrl AS VARIANT PTR, BYVAL Flags AS VARIANT PTR = NULL, BYVAL TargetFrameName AS VARIANT PTR = NULL, BYVAL PostData AS VARIANT PTR = NULL, BYVAL Headers AS VARIANT PTR = NULL) AS HRESULT
   CWB_DP("CWebBrowser.Navigate - vUrl")
   IF m_pWebBrowser THEN FUNCTION = m_pWebBrowser->Navigate2(vUrl, Flags, TargetFrameName, PostData, Headers)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Establishes a connection between the connection point object and the client's sink.
' ========================================================================================
PRIVATE FUNCTION CWebBrowser.Advise () AS HRESULT
   CWB_DP("CWebBrowser.Advise")
   ' // Create an instance of our CWebBrowserEventsImpl class and connect the WebBrowser events
   m_pEvtObj = NEW CWebBrowserEventsImpl(m_pOleCon->hWindow)
   m_pEvtObj->m_pWebBrowser = m_pWebBrowser
   RETURN m_pOleCon->Advise(m_pEvtObj, DIID_DWebBrowserEvents2)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Releases the events connection.
' ========================================================================================
PRIVATE FUNCTION CWebBrowser.Unadvise () AS HRESULT
   CWB_DP("CWebBrowser.Unadvise")
   RETURN m_pOleCon->Unadvise
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets our implementation of the IDocHostUIHandler interface to customize the WebBrowser.
' ========================================================================================
PRIVATE FUNCTION CWebBrowser.SetUIHandler () AS HRESULT
   CWB_DP("-BEGIN CWebBrowser.SetUIHandler")

   ' // After the MSHTML document has been loaded, we retrieve a reference to its
   ' // ICustomDoc interface and give him a pointer to our IDocHostUIHandler interface
   ' // to allow for customization.
   IF m_pWebBrowser = NULL THEN RETURN E_POINTER

   ' // Get a reference to the active document
   DIM pIHTMLDocument2 AS IHTMLDocument2 PTR
   m_pWebBrowser->get_Document(@cast(IDispatch PTR, pIHTMLDocument2))
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
   IF m_pDocHostUIHandler = NULL THEN m_pDocHostUIHandler = NEW CDocHostUIHandler2Impl(m_pOleCon->hWindow, m_pWebBrowser)
   ' // Set our IDocHostUIHandler interface for MSHTML
   ' // MSHTML will release its previous IDocHostUIHandler interface
   ' // (if one is present) and call pDocHostUIHandler's AddRef method.
   IF m_pDocHostUIHandler THEN pICustomDoc->lpvtbl->SetUIHandler(pICustomDoc, _
      cast(IDocHostUIHandler PTR, cast(ULONG_PTR, m_pDocHostUIHandler)))
   IUnknown_Release(pICustomDoc)

   CWB_DP("-END CWebBrowser.SetUIHandler")
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets pointers to user implemented callback procedures.
' ========================================================================================
PRIVATE FUNCTION CWebBrowser.SetEventProc (BYVAL pwszEventName AS WSTRING PTR, BYVAL pProc AS ANY PTR) AS HRESULT
   CWB_DP("CWebBrowser.SetEventHandler")
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
PRIVATE FUNCTION CWebBrowser.SetUIEventProc (BYVAL pwszEventName AS WSTRING PTR, BYVAL pProc AS ANY PTR) AS HRESULT
   CWB_DP("CWebBrowser.SetUIEventHandler")
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
PRIVATE FUNCTION CWebBrowser.WriteHtml (BYREF cwsHtml AS CWSTR, BYVAL cr AS BOOLEAN = FALSE) AS HRESULT
   RETURN AfxWriteHtml(m_pWebBrowser, cwsHtml, cr)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the ID of the active element (the object that has the focus when the parent
' document has focus).
' Return value: The ID of the active element.
' ========================================================================================
PRIVATE FUNCTION CWebBrowser.GetActiveElementId () AS CWSTR
   RETURN AfxGetActiveElementId(m_pWebBrowser)
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
PRIVATE FUNCTION CWebBrowser.GetElementValueById (BYREF cwsId AS CWSTR) AS VARIANT
   RETURN AfxGetElementValueById(m_pWebBrowser, cwsId)
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
PRIVATE FUNCTION CWebBrowser.SetElementValueById (BYREF cwsId AS CWSTR, BYVAL vValue AS VARIANT) AS HRESULT
   RETURN AfxSetElementValueById(m_pWebBrowser, cwsId, vValue)
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
PRIVATE FUNCTION CWebBrowser.GetElementInnerHtmlById (BYREF cwsId AS CWSTR) AS CWSTR
   RETURN AfxGetElementInnerHtmlById(m_pWebBrowser, cwsId)
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
PRIVATE FUNCTION CWebBrowser.SetElementInnerHtmlById (BYREF cwsId AS CWSTR, BYREF cwsHtml AS CWSTR) AS HRESULT
   RETURN AfxSetElementInnerHtmlById(m_pWebBrowser, cwsId, cwsHtml)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the focus in the specified element.
' Parameter:
' - cbsId = The identifier.
' Return value:
' - S_OK if successful, or an error value otherwise.
' ========================================================================================
PRIVATE FUNCTION CWebBrowser.SetElementFocusById (BYREF cwsId AS CWSTR) AS HRESULT
   RETURN AfxSetElementFocusById(m_pWebBrowser, cwsId)
END FUNCTION
' ========================================================================================

' ========================================================================================
'                  *** IWebBrowser2 interface additional methods ***
' ========================================================================================

' ========================================================================================
' Navigates backward one item in the history list.
' ========================================================================================
PRIVATE FUNCTION CWebBrowser.GoBack () AS HRESULT
   IF m_pWebBrowser THEN RETURN m_pWebBrowser->GoBack
END FUNCTION
' ========================================================================================
' ========================================================================================
' Navigates backward one item in the history list.
' ========================================================================================
PRIVATE FUNCTION CWebBrowser.GoForward () AS HRESULT
   IF m_pWebBrowser THEN RETURN m_pWebBrowser->GoForward
END FUNCTION
' ========================================================================================
' ========================================================================================
' Navigates to the current home or start page.
' ========================================================================================
PRIVATE FUNCTION CWebBrowser.GoHome () AS HRESULT
   IF m_pWebBrowser THEN RETURN m_pWebBrowser->GoHome
END FUNCTION
' ========================================================================================
' ========================================================================================
' Navigates to the current home or start page.
' ========================================================================================
PRIVATE FUNCTION CWebBrowser.GoSearch () AS HRESULT
   IF m_pWebBrowser THEN RETURN m_pWebBrowser->GoSearch
END FUNCTION
' ========================================================================================
' ========================================================================================
' Reloads the file that is currently displayed in the object.
' ========================================================================================
PRIVATE FUNCTION CWebBrowser.Refresh () AS HRESULT
   IF m_pWebBrowser THEN RETURN m_pWebBrowser->Refresh
END FUNCTION
' ========================================================================================
' ========================================================================================
' Reloads the file that is currently displayed with the specified refresh level.
' ========================================================================================
PRIVATE FUNCTION CWebBrowser.Refresh2 (BYVAL nLevel AS RefreshConstants) AS HRESULT
   DIM vLevel AS VARIANT
   vLevel.vt = VT_I4 : vLevel.lVal = nLevel
   IF m_pWebBrowser THEN RETURN m_pWebBrowser->Refresh2(@vLevel)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Cancels a pending navigation or download, and stops dynamic page elements, such as
' background sounds and animations.
' ========================================================================================
PRIVATE FUNCTION CWebBrowser.Stop () AS HRESULT
   IF m_pWebBrowser THEN RETURN m_pWebBrowser->Stop
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets the automation object of the active document, if any.
' ========================================================================================
PRIVATE PROPERTY CWebBrowser.Document () AS IHtmlDocument2 PTR
   DIM pdisp AS IDispatch PTR
   IF m_pWebBrowser THEN m_pWebBrowser->get_Document(@pdisp)
   PROPERTY = cast(IHtmlDocument2 PTR, pdisp)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Retrieves the path or title of the resource that is currently displayed.
' ========================================================================================
PRIVATE PROPERTY CWebBrowser.LocationName () AS CWSTR
   DIM bstrName AS AFX_BSTR
   IF m_pWebBrowser THEN m_pWebBrowser->get_LocationName(@bstrName)
   PROPERTY = *bstrName
   SysFreestring bstrName
END PROPERTY
' ========================================================================================
' ========================================================================================
' Retrieves the URL of the resource that is currently displayed.
' ========================================================================================
PRIVATE PROPERTY CWebBrowser.LocationURL () AS CWSTR
   DIM bstrName AS AFX_BSTR
   IF m_pWebBrowser THEN m_pWebBrowser->get_LocationURL(@bstrName)
   PROPERTY = *bstrName
   SysFreestring bstrName
END PROPERTY
' ========================================================================================
' ========================================================================================
' TRUE if the object is engaged in a navigation or downloading operation; FALSE, otherwise.
' ========================================================================================
PRIVATE PROPERTY CWebBrowser.Busy () AS BOOLEAN
   DIM bBusy AS VARIANT_BOOL
   IF m_pWebBrowser THEN m_pWebBrowser->get_Busy(@bBusy)
   PROPERTY = bBusy
END PROPERTY
' ========================================================================================
' ========================================================================================
' Queries the object for the status of commands using the IOleCommandTarget::QueryStatus method.
' ========================================================================================
PRIVATE FUNCTION CWebBrowser.QueryStatusWB (BYVAL cmdID AS OLECMDID) AS OLECMDF
   DIM cmdf AS OLECMDF
   IF m_pWebBrowser THEN m_pWebBrowser->QueryStatusWB(cmdID, @cmdf)
   RETURN cmdf
END FUNCTION
' ========================================================================================
' ========================================================================================
' Executes a command and returns the status of the command execution using the IOleCommandTarget interface.
' ========================================================================================
PRIVATE FUNCTION CWebBrowser.ExecWB (BYVAL cmdID AS OLECMDID, BYVAL cmdexecopt AS OLECMDEXECOPT, BYVAL pvaIn AS VARIANT PTR = NULL, BYVAL pvaOut AS VARIANT PTR = NULL) AS HRESULT
   IF m_pWebBrowser THEN RETURN(m_pWebBrowser->ExecWB(cmdID, cmdexecopt, pvaIn, pvaOut))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets the ready state of the object.
' ========================================================================================
PRIVATE PROPERTY CWebBrowser.ReadyState () AS tagREADYSTATE
   DIM nState AS tagREADYSTATE
   IF m_pWebBrowser THEN m_pWebBrowser->get_ReadyState(@nState)
   PROPERTY = nState
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets or gets a value that indicates whether the object is registered as a top-level browser window.
' ========================================================================================
PRIVATE PROPERTY CWebBrowser.RegisterAsBrowser () AS BOOLEAN
   DIM bRegister AS VARIANT_BOOL
   IF m_pWebBrowser THEN m_pWebBrowser->get_RegisterAsBrowser(@bRegister)
   PROPERTY = bRegister
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CWebBrowser.RegisterAsBrowser (BYVAL bRegister AS BOOLEAN)
   IF m_pWebBrowser THEN m_pWebBrowser->put_RegisterAsBrowser(bRegister)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Sets or gets a value that indicates whether the object is registered as a drop target for navigation.
' ========================================================================================
PRIVATE PROPERTY CWebBrowser.RegisterAsDropTarget () AS BOOLEAN
   DIM bRegister AS VARIANT_BOOL
   IF m_pWebBrowser THEN m_pWebBrowser->get_RegisterAsDropTarget(@bRegister)
   PROPERTY = bRegister
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CWebBrowser.RegisterAsDropTarget (BYVAL bRegister AS BOOLEAN)
   IF m_pWebBrowser THEN m_pWebBrowser->put_RegisterAsDropTarget(bRegister)
END PROPERTY
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
   DIM wszClassName AS WSTRING * 260
   DIM hwndChild AS .HWND = AfxGetControlHandle(hwnd, cID)
   GetClassNameW hwndChild, wszClassName, SIZEOF(wszClassName)
   IF wszClassName <> OC_CLASSNAME THEN RETURN NULL
   DIM pOleCon AS COleCon PTR = CAST(COleCon PTR, .GetWindowLongPtrW(hwndChild, 0))
   IF pOleCon THEN RETURN cast(Afx_IWebBrowser2 PTR, cast(ULONG_PTR, pOleCon->m_pData->m_pOcx))
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
   DIM hr AS HRESULT = pWebBrowser->get_Document(@cast(IDispatch PTR, pIHTMLDocument2))
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
   DIM hr AS HRESULT = pWebBrowser->get_Document(@cast(IDispatch PTR, pIHTMLDocument2))
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
   DIM hr AS HRESULT = pWebBrowser->get_Document(@cast(IDispatch PTR, pIHTMLDocument2))
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
   DIM hr AS HRESULT = pWebBrowser->get_Document(@cast(IDispatch PTR, pIHTMLDocument2))
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
   DIM hr AS HRESULT = pWebBrowser->get_Document(@cast(IDispatch PTR, pIHTMLDocument2))
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
   DIM hr AS HRESULT = pWebBrowser->get_Document(@cast(IDispatch PTR, pIHTMLDocument2))
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
   DIM hr AS HRESULT = pWebBrowser->get_Document(@cast(IDispatch PTR, pIHTMLDocument2))
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

END NAMESPACE
