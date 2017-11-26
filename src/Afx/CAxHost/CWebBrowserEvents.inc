' ########################################################################################
' CWebBrowserEvents class
' Implementation of the WebBrowser events sink class
' ########################################################################################
TYPE CWebBrowserEvents EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULONG
   DECLARE VIRTUAL FUNCTION GetTypeInfoCount (BYVAL pctinfo AS UINT PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION GetTypeInfo (BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION GetIDsOfNames (BYVAL riid AS CONST IID CONST PTR, BYVAL rgszNames AS LPOLESTR PTR, BYVAL cNames AS UINT, BYVAL lcid AS LCID, BYVAL rgDispId AS DISPID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION Invoke (BYVAL dispIdMember AS DISPID, BYVAL riid AS CONST IID CONST PTR, BYVAL lcid AS LCID, BYVAL wFlags AS WORD, BYVAL pDispParams AS DISPPARAMS PTR, BYVAL pVarResult AS VARIANT PTR, BYVAL pExcepInfo AS EXCEPINFO PTR, BYVAL puArgErr AS UINT PTR) AS HRESULT

   DECLARE CONSTRUCTOR (BYVAL pWebCtx AS CWebCtx PTR)
   DECLARE DESTRUCTOR

   m_pWebCtx AS CWebCtx PTR                             ' // Pointer to the CWebtx class
   m_pWebBrowser AS Afx_IWebBrowser2 PTR                ' // Pointer to the IWebBrowser interface
   m_pHTMLDocument2 AS IHTMLDocument2 PTR               ' // Pointer to the IHTMLDocument interface
   m_pHTMLDocumentEvents2 AS CHTMLDocumentEvents2 PTR   ' // Pointer to the CHTMLDocumentEvents2Impl class
   m_dwDocEventsCookie AS DWORD                         ' // Events cookie
   cRef AS ULONG                                        ' // Reference count

   ' // Pointers to user implemented event procedures
   m_StatusTextChangeProc AS SUB (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pwszText AS WSTRING PTR)
   m_DownloadCompleteProc AS SUB (BYVAL pWebCtx AS CWebCtx PTR)
   m_CommandStateChangeProc AS SUB (BYVAL pWebCtx AS CWebCtx PTR, BYVAL nCommand AS LONG, BYVAL fEnable AS VARIANT_BOOL)
   m_DownloadBeginProc AS SUB (BYVAL pWebCtx AS CWebCtx PTR)
   m_ProgressChangeProc AS SUB (BYVAL pWebCtx AS CWebCtx PTR, BYVAL Progress AS LONG, BYVAL ProgressMax AS LONG)
   m_PropertyChangeProc AS SUB (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pwszProperty AS WSTRING PTR)
   m_TitleChangeProc AS SUB (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pwszText AS WSTRING PTR)
   m_PrintTemplateInstantiationProc AS SUB (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pdisp AS IDispatch PTR)
   m_PrintTemplateTeardownProc AS SUB (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pdisp AS IDispatch PTR)
   m_BeforeNavigate2Proc AS SUB (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pdisp AS IDispatch PTR, _
      BYVAL vUrl AS VARIANT PTR, BYVAL Flags AS VARIANT PTR, BYVAL TargetFrameName AS VARIANT PTR, _
      BYVAL PostData AS VARIANT PTR, BYVAL Headers AS VARIANT PTR, BYVAL pbCancel AS VARIANT_BOOL PTR)
   m_NavigateComplete2Proc AS SUB (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pdisp AS IDispatch PTR, BYVAL vUrl AS VARIANT PTR)
   m_OnVisibleProc AS SUB (BYVAL pWebCtx AS CWebCtx PTR, BYVAL bVisible AS VARIANT_BOOL)
   m_DocumentCompleteProc AS SUB (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pdisp AS IDispatch PTR, BYVAL vUrl AS VARIANT PTR)
   m_WindowSetResizableProc AS SUB (BYVAL pWebCtx AS CWebCtx PTR, BYVAL bResizable AS VARIANT_BOOL)
   m_WindowClosingProc AS SUB (BYVAL pWebCtx AS CWebCtx PTR, BYVAL IsChildWindow AS VARIANT_BOOL, BYVAL pbCancel AS VARIANT_BOOL PTR)
   m_WindowSetLeftProc AS SUB (BYVAL pWebCtx AS CWebCtx PTR, BYVAL nLenft AS LONG)
   m_WindowSetTopProc AS SUB (BYVAL pWebCtx AS CWebCtx PTR, BYVAL nTop AS LONG)
   m_WindowSetWidthProc AS SUB (BYVAL pWebCtx AS CWebCtx PTR, BYVAL nWidth AS LONG)
   m_WindowSetHeightProc AS SUB (BYVAL pWebCtx AS CWebCtx PTR, BYVAL nHeight AS LONG)
   m_ClientToHostWindowProc AS SUB (BYVAL pWebCtx AS CWebCtx PTR, BYVAL CX AS LONG PTR, BYVAL CY AS LONG PTR)
   m_SetSecureLockIconProc AS SUB (BYVAL pWebCtx AS CWebCtx PTR, BYVAL SecureLockIcon AS LONG)
   m_FileDownloadProc AS SUB (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pbCancel AS VARIANT_BOOL PTR)
   m_NavigateErrorProc AS SUB (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pdisp AS IDispatch PTR, _
      BYVAL vUrl AS VARIANT PTR, BYVAL vFrame AS VARIANT PTR, BYVAL StatusCode AS VARIANT PTR, BYVAL pbCancel AS VARIANT_BOOL PTR)
   m_PrivacyImpactedStateChangeProc AS SUB (BYVAL pWebCtx AS CWebCtx PTR, BYVAL bImpacted AS VARIANT_BOOL)
   m_NewWindow2Proc AS SUB (BYVAL pWebCtx AS CWebCtx PTR, BYVAL ppdispVal AS IDispatch PTR PTR, BYVAL pbCancel AS VARIANT_BOOL PTR)
   m_NewWindow3Proc AS SUB (BYVAL pWebCtx AS CWebCtx PTR, BYVAL ppdispVal AS IDispatch PTR PTR, _
      BYVAL pbCancel AS VARIANT_BOOL PTR, BYVAL dwFlags AS ULONG, BYVAL pwszUrlContext AS WSTRING PTR, BYVAL pwszUrl AS WSTRING PTR)
   m_WindowStateChangedProc AS SUB (BYVAL pWebCtx AS CWebCtx PTR, BYVAL dwFlags AS LONG, BYVAL dwValidFlagsMask AS LONG)

END TYPE
' ########################################################################################

' ========================================================================================
' Constructors
' ========================================================================================
CONSTRUCTOR CWebBrowserEvents (BYVAL pWebCtx AS CWebCtx PTR)
   CWBX_DP("CONSTRUCTOR CWebBrowserEvents")
   ' // Create an instance of the CHTMLDocumentEvents2 class
   m_pWebCtx = pWebCtx
   m_pHTMLDocumentEvents2 = NEW CHTMLDocumentEvents2(m_pWebCtx, m_pWebBrowser)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destructor
' ========================================================================================
DESTRUCTOR CWebBrowserEvents
   CWBX_DP("DESTRUCTOR CWebBrowserEvents")
   ' // Delete the HTMLDocumentEvents2 class
   IF m_pHTMLDocument2 THEN m_pHTMLDocument2->lpvtbl->Release(m_pHTMLDocument2)
   IF m_pHTMLDocumentEvents2 THEN DELETE m_pHTMLDocumentEvents2
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns a pointer to our implementation of the WebBrowser events class.
' ========================================================================================
FUNCTION CWebBrowserEvents.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   CWBX_DP("CWebBrowserEvents.QueryInterface")
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   IF IsEqualIID(riid, @DIID_DWebBrowserEvents2) OR _
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
' This is not a COM class, so the life of the object is not managed by a reference count.
' ========================================================================================
FUNCTION CWebBrowserEvents.AddRef () AS ULONG
   this.cRef += 1
   CWBX_DP("CWebBrowserEvents AddRef = " & WSTR(this.cRef))
   RETURN this.cRef
END FUNCTION
' ========================================================================================

' ========================================================================================
' This is not a COM class, so the life of the object is not managed by a reference count.
' ========================================================================================
FUNCTION CWebBrowserEvents.Release () AS ULONG
   this.cRef -= 1
   CWBX_DP("CWebBrowserEvents Release = " & WSTR(this.cRef))
   FUNCTION = this.cRef
   ' // Don't delete it because this is not a COM class
   'IF this.cRef = 0 THEN DELETE @this
END FUNCTION
' ========================================================================================

' ========================================================================================
' The number of type information interfaces provided by the object. If the object provides
' type information, this number is 1; otherwise the number is 0.
' ========================================================================================
FUNCTION CWebBrowserEvents.GetTypeInfoCount (BYVAL pctInfo AS UINT PTR) AS HRESULT
   CWBX_DP("CWebBrowserEvents.GetTypeInfoCount")
   *pctInfo = 0
   RETURN E_NOTIMPL
END FUNCTION
' ========================================================================================

' ========================================================================================
' The type information for an object, which can then be used to get the type information
' for an interface.
' ========================================================================================
FUNCTION CWebBrowserEvents.GetTypeInfo (BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   CWBX_DP("CWebBrowserEvents.GetTypeInfo")
   RETURN E_NOTIMPL
END FUNCTION
' ========================================================================================

' ========================================================================================
' Maps a single member and an optional set of argument names to a corresponding set of
' integer DISPIDs, which can be used on subsequent calls to Invoke.
' ========================================================================================
FUNCTION CWebBrowserEvents.GetIDsOfNames (BYVAL riid AS CONST IID CONST PTR, BYVAL rgszNames AS LPOLESTR PTR, BYVAL cNames AS UINT, BYVAL lcid AS LCID, BYVAL rgDispId AS DISPID PTR) AS HRESULT
   CWBX_DP("CWebBrowserEvents.GetIDsOfNames")
   RETURN E_NOTIMPL
END FUNCTION
' ========================================================================================

' ========================================================================================
' Provides access to properties and methods exposed by an object.
' Parameters in the rgvarg array of variants of the DISPPARAMS structures are zero based
' and in reverse order.
' ========================================================================================
FUNCTION CWebBrowserEvents.Invoke (BYVAL dispIdMember AS DISPID, BYVAL riid AS CONST IID CONST PTR, BYVAL lcid AS LCID, BYVAL wFlags AS WORD, BYVAL pDispParams AS DISPPARAMS PTR, BYVAL pVarResult AS VARIANT PTR, BYVAL pExcepInfo AS EXCEPINFO PTR, BYVAL puArgErr AS UINT PTR) AS HRESULT
   CWBX_DP("CWebBrowserEvents.Invoke dispid = " & WSTR(dispIdMember))

   SELECT CASE dispIdMember

      ' ----------------------------------------------------------------------------------
      CASE 102   ' // StatusTextChange
      ' ----------------------------------------------------------------------------------
      CWBX_DP("CWebBrowserEvents.StatusTextChange")
      ' // Fires when the status bar text of the object has changed.
      ' // You can use the information provided by this event to update the text of a status bar.
      ' // Remarks: Do not free the BSTR (it is owned by the caller).
      IF m_StatusTextChangeProc THEN m_StatusTextChangeProc(m_pWebCtx, pDispParams->rgvarg[0].bstrVal)

      ' ----------------------------------------------------------------------------------
      CASE 104   ' // DownloadComplete
      ' ----------------------------------------------------------------------------------
      CWBX_DP("CWebBrowserEvents.DownloadComplete")
      ' // Fires when a navigation operation finishes, is halted, or fails.
      IF m_DownloadCompleteProc THEN m_DownloadCompleteProc(m_pWebCtx)

      ' ----------------------------------------------------------------------------------
      CASE 105   ' // CommandStateChange
      ' ----------------------------------------------------------------------------------
      CWBX_DP("CWebBrowserEvents.CommandStateChange")
      ' // Fires when the enabled state of a command changes.
      IF m_CommandStateChangeProc THEN m_CommandStateChangeProc(m_pWebCtx, pDispParams->rgvarg[1].lVal, pDispParams->rgvarg[0].boolVal)

      ' ----------------------------------------------------------------------------------
      CASE 106   ' // DownloadBegin
      ' ----------------------------------------------------------------------------------
      CWBX_DP("CWebBrowserEvents.DownloadBegin")
      ' // Fires when a navigation operation is beginning.
      IF m_DownloadBeginProc THEN m_DownloadBeginProc(m_pWebCtx)

      ' ----------------------------------------------------------------------------------
      CASE 108   ' // ProgressChange
      ' ----------------------------------------------------------------------------------
      CWBX_DP("CWebBrowserEvents.ProgressChange")
      ' // Fires when the progress of a download operation is updated on the object.
      ' // The container can use the information provided by this event to display the number of bytes
      ' // downloaded so far or to update a progress indicator.
      ' // To calculate the percentage of progress to show in a progress indicator, multiply the value
      ' // of Progress by 100 and divide by the value of ProgressMax (unless Progress is -1, in which
      ' // case the container can indicate that the operation is finished or hide the progress indicator).
      IF m_ProgressChangeProc THEN m_ProgressChangeProc(m_pWebCtx, pDispParams->rgvarg[1].lVal, pDispParams->rgvarg[0].lVal)

      ' ----------------------------------------------------------------------------------
      CASE 112   ' // PropertyChange
      ' ----------------------------------------------------------------------------------
      CWBX_DP("CWebBrowserEvents.PropertyChange")
      ' // Fires when the PutProperty method of the object changes the value of a property.
      IF m_PropertyChangeProc THEN m_PropertyChangeProc(m_pWebCtx, pDispParams->rgvarg[0].bstrVal)

      ' ----------------------------------------------------------------------------------
      CASE 113   ' // TitleChange
      ' ----------------------------------------------------------------------------------
      CWBX_DP("CWebBrowserEvents.TitleChange")
      ' // Fires when the title of a document in the object becomes available or changes.
      ' // Because the title might change while an HTML page is downloading, the URL of the
      ' // document is set as the title. After the title specified in the HTML page, if
      ' // there is one, is parsed, the title is changed to reflect the actual title.
      IF m_TitleChangeProc THEN m_TitleChangeProc(m_pWebCtx, pDispParams->rgvarg[0].bstrVal)

      ' ----------------------------------------------------------------------------------
      CASE 225   ' // PrintTemplateInstantiation
      ' ----------------------------------------------------------------------------------
      CWBX_DP("CWebBrowserEvents.PrintTemplateInstantiation")
      ' // Fires when a print template has been instantiated.
      IF m_PrintTemplateInstantiationProc THEN m_PrintTemplateInstantiationProc(m_pWebCtx, pDispParams->rgvarg[0].pdispVal)

      ' ----------------------------------------------------------------------------------
      CASE 226   ' // PrintTemplateTeardown
      ' ----------------------------------------------------------------------------------
      CWBX_DP("CWebBrowserEvents.PrintTemplateTeardown")
      ' // Fires when a print template has been destroyed.
      IF m_PrintTemplateTeardownProc THEN m_PrintTemplateTeardownProc(m_pWebCtx, pDispParams->rgvarg[0].pdispVal)

      ' ----------------------------------------------------------------------------------
      CASE 227   ' // UpdatePageStatus
      ' ----------------------------------------------------------------------------------
      CWBX_DP("CWebBrowserEvents.UpdatePageStatus")
      ' // Not implemented.

      ' ----------------------------------------------------------------------------------
      CASE 250   ' // BeforeNavigate2
      ' ----------------------------------------------------------------------------------
      CWBX_DP("CWebBrowserEvents.BeforeNavigate2")
      ' // Fires before navigation occurs in the given object (on either a window or frameset element).
      IF m_BeforeNavigate2Proc THEN m_BeforeNavigate2Proc(m_pWebCtx, pDispParams->rgvarg[6].pdispVal,_
         pDispParams->rgvarg[5].pvarVal, pDispParams->rgvarg[4].pvarVal, pDispParams->rgvarg[3].pvarVal, _
         pDispParams->rgvarg[2].pvarVal, pDispParams->rgvarg[1].pvarVal, pDispParams->rgvarg[0].pboolVal)

      ' ----------------------------------------------------------------------------------
      CASE 252   ' // NavigateComplete2
      ' ----------------------------------------------------------------------------------
      CWBX_DP("CWebBrowserEvents.NavigateComplete2")
      ' // Fires after a navigation to a link is completed on either a window or frameset element.
      IF m_NavigateComplete2Proc THEN m_NavigateComplete2Proc(m_pWebCtx, pDispParams->rgvarg[1].pdispVal, @pDispParams->rgvarg[0])

      ' ----------------------------------------------------------------------------------
      CASE 253   ' // OnQuit
      ' ----------------------------------------------------------------------------------
      CWBX_DP("CWebBrowserEvents.OnQuit")
      ' // Fires before the Windows Internet Explorer application quits.
      ' // The WebBrowser object ignores this event.

      ' ----------------------------------------------------------------------------------
      CASE 254   ' // OnVisible
      ' ----------------------------------------------------------------------------------
      CWBX_DP("CWebBrowserEvents.OnVisible")
      ' //  Fires when the IWebBrowser2.Visible property of the object is changed.
      IF m_OnVisibleProc THEN m_OnVisibleProc(m_pWebCtx, pDispParams->rgvarg[0].boolVal)

      ' ----------------------------------------------------------------------------------
      CASE 259   ' // DocumentComplete
      ' ----------------------------------------------------------------------------------
      CWBX_DP("CWebBrowserEvents.DocumentComplete")
      ' // Fires when a document is completely loaded and initialized.
      IF m_DocumentCompleteProc = NULL THEN RETURN S_OK
      IF m_DocumentCompleteProc THEN m_DocumentCompleteProc(m_pWebCtx, pDispParams->rgvarg[1].pdispVal, @pDispParams->rgvarg[0])
      ' // Exit if the user has not set a callback for the Hmtl document events
      IF m_pHTMLDocumentEvents2 THEN
         IF m_pHTMLDocumentEvents2->m_pHtmlDocumentEventsProc = NULL THEN RETURN S_OK
      END IF
      ' // --- Connect/disconnect with events fired by the web page ----------------------
      DIM pWebBrowser2 AS Afx_IWebBrowser2 PTR = cast(Afx_IWebBrowser2 PTR, cast(ULONG_PTR, pDispParams->rgvarg[1].pdispVal))
      IF pWebBrowser2 = NULL THEN EXIT FUNCTION
      ' extern DIID_HTMLDocumentEvents2 is broken and causes the linker to fail
      DIM DIID_HTMLDocumentEvents2_ AS GUID = (&h3050F613, &h98B5, &h11CF, {&hBB, &h82, &h00, &hAA, &h00, &hBD, &hCE, &h0B})
      ' // If there was a previous loaded page, disconnect from the events
      IF m_pHTMLDocument2 <> NULL AND m_dwDocEventsCookie <> 0 THEN
         IF AfxUnadvise(m_pHTMLDocument2, DIID_HTMLDocumentEvents2_, m_dwDocEventsCookie) = S_OK THEN
            m_pHTMLDocument2->lpvtbl->Release(m_pHTMLDocument2)
            m_pHTMLDocument2 = NULL
            m_dwDocEventsCookie = 0
         END IF
      END IF
      IF m_pHTMLDocumentEvents2 <> NULL AND m_dwDocEventsCookie = 0 THEN
         ' // Connect with the HTMLDocumentEvents2 interface
         pWebBrowser2->get_Document(@cast(ANY PTR, m_pHTMLDocument2))
         IF m_pHTMLDocument2 = NULL THEN EXIT FUNCTION
         AfxAdvise(m_pHTMLDocument2, m_pHTMLDocumentEvents2, DIID_HTMLDocumentEvents2_, @m_dwDocEventsCookie)
      END IF
      ' // ---------------------------------------------------------s----------------------

      ' ----------------------------------------------------------------------------------
      CASE 262   ' // WindowSetResizable
      ' ----------------------------------------------------------------------------------
      CWBX_DP("CWebBrowserEvents.WindowSetResizable")
      ' // Fires to indicate whether the host window should allow or disallow resizing of the object.
      IF m_WindowSetResizableProc THEN m_WindowSetResizableProc(m_pWebCtx, pDispParams->rgvarg[0].boolVal)

      ' ----------------------------------------------------------------------------------
      CASE 263   ' // WindowClosing
      ' ----------------------------------------------------------------------------------
      CWBX_DP("CWebBrowserEvents.WindowClosing")
      ' // Fires when the window of the object is about to be closed by script.
      IF m_WindowClosingProc THEN m_WindowClosingProc(m_pWebCtx, pDispParams->rgvarg[1].boolVal, pDispParams->rgvarg[0].pboolVal)

      ' ----------------------------------------------------------------------------------
      CASE 264   ' // WindowSetLeft
      ' ----------------------------------------------------------------------------------
      CWBX_DP("CWebBrowserEvents.WindowSetLeft")
      ' // Fires when the object changes its left position.
      IF m_WindowSetLeftProc THEN m_WindowSetLeftProc(m_pWebCtx, pDispParams->rgvarg[0].lVal)

      ' ----------------------------------------------------------------------------------
      CASE 265   ' // WindowSetTop
      ' ----------------------------------------------------------------------------------
      CWBX_DP("CWebBrowserEvents.WindowSetTop")
      ' // Fires when the object changes its top position.
      IF m_WindowSetTopProc THEN m_WindowSetTopProc(m_pWebCtx, pDispParams->rgvarg[0].lVal)

      ' ----------------------------------------------------------------------------------
      CASE 266   ' // WindowSetWidth
      ' ----------------------------------------------------------------------------------
      CWBX_DP("CWebBrowserEvents.WindowSetWidth")
      ' // Fires when the object changes its width.
      IF m_WindowSetWidthProc THEN m_WindowSetWidthProc(m_pWebCtx, pDispParams->rgvarg[0].lVal)

      ' ----------------------------------------------------------------------------------
      CASE 267   ' // WindowSetHeight
      ' ----------------------------------------------------------------------------------
      CWBX_DP("CWebBrowserEvents.WindowSetHeight")
      ' // Fires when the object changes its height.
      IF m_WindowSetHeightProc THEN m_WindowSetHeightProc(m_pWebCtx, pDispParams->rgvarg[0].lVal)

      ' ----------------------------------------------------------------------------------
      CASE 268   ' // ClientToHostWindow
      ' ----------------------------------------------------------------------------------
      CWBX_DP("CWebBrowserEvents.ClientToHostWindow")
      ' // Fires to request that the client window size be converted to the host window size.
      IF m_ClientToHostWindowProc THEN m_ClientToHostWindowProc(m_pWebCtx, pDispParams->rgvarg[1].plVal, pDispParams->rgvarg[0].plVal)

      ' ----------------------------------------------------------------------------------
      CASE 269   ' // SetSecureLockIcon
      ' ----------------------------------------------------------------------------------
      CWBX_DP("CWebBrowserEvents.SetSecureLockIcon")
      ' // Fires when there is a change in encryption level.
      IF m_SetSecureLockIconProc THEN m_SetSecureLockIconProc(m_pWebCtx, pDispParams->rgvarg[0].lVal)

      ' ----------------------------------------------------------------------------------
      CASE 270   ' // FileDownload
      ' ----------------------------------------------------------------------------------
      CWBX_DP("CWebBrowserEvents.FileDownload")
      ' // Fires to indicate that a file download is about to occur. If a file download
      ' // dialog is to be displayed, this event is fired prior to the display of the dialog.
      IF m_FileDownloadProc THEN m_FileDownloadProc(m_pWebCtx, pDispParams->rgvarg[0].pboolVal)

      ' ----------------------------------------------------------------------------------
      CASE 271   ' // NavigateError
      ' ----------------------------------------------------------------------------------
      CWBX_DP("CWebBrowserEvents.NavigateError")
      ' // Fires when an error occurs during navigation.
      IF m_NavigateErrorProc THEN m_NavigateErrorProc(m_pWebCtx, pDispParams->rgvarg[4].pdispVal, _
         @pDispParams->rgvarg[3], @pDispParams->rgvarg[2], @pDispParams->rgvarg[1], pDispParams->rgvarg[0].pboolVal)

      ' ----------------------------------------------------------------------------------
      CASE 272   ' // PrivacyImpactedStateChange
      ' ----------------------------------------------------------------------------------
      CWBX_DP("CWebBrowserEvents.PrivacyImpactedStateChange")
      ' // Fired when an event occurs that impacts privacy or when a user navigates away
      ' // from a URL that has impacted privacy.
      ' DIM bImpacted AS BOOLEAN = pDispParams->rgvarg[0].boolVal
      IF m_PrivacyImpactedStateChangeProc THEN m_PrivacyImpactedStateChangeProc(m_pWebCtx, _
         pDispParams->rgvarg[0].boolVal)

      ' ----------------------------------------------------------------------------------
      CASE 251   ' // NewWindow2
      ' ----------------------------------------------------------------------------------
      CWBX_DP("CWebBrowserEvents.NewWindow2")
      ' // Raised when a new window is to be created.
      IF m_NewWindow2Proc THEN m_NewWindow2Proc(m_pWebCtx, pDispParams->rgvarg[1].ppdispVal, _
         pDispParams->rgvarg[0].pboolVal)

      ' ----------------------------------------------------------------------------------
      CASE 273   ' // NewWindow3
      ' ----------------------------------------------------------------------------------
      CWBX_DP("CWebBrowserEvents.NewWindow3")
      ' // Raised when a new window is to be created.
      IF m_NewWindow3Proc THEN m_NewWindow3Proc(m_pWebCtx, pDispParams->rgvarg[4].ppdispVal, _
         pDispParams->rgvarg[3].pboolVal, pDispParams->rgvarg[2].ulVal, pDispParams->rgvarg[1].bstrVal, _
         pDispParams->rgvarg[0].bstrVal)

      ' ----------------------------------------------------------------------------------
      CASE 283   ' // WindowStateChanged
      ' ----------------------------------------------------------------------------------
      CWBX_DP("CWebBrowserEvents.WindowStateChanged")
      ' // Fires when the visibility state of a content window, such as the browser window
      ' // or a tab, changes.
      IF m_WindowStateChangedProc THEN m_WindowStateChangedProc(m_pWebCtx, pDispParams->rgvarg[1].lVal, pDispParams->rgvarg[0].lVal)

   END SELECT

   RETURN S_OK

END FUNCTION
' ========================================================================================
