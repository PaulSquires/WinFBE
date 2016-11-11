' ########################################################################################
' CWebBrowserEventsImpl class
' Implementation of the WebBrowser events sink class
' ########################################################################################
TYPE CWebBrowserEventsImpl EXTENDS Afx_IDispatch

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULONG
   DECLARE VIRTUAL FUNCTION GetTypeInfoCount (BYVAL pctinfo AS UINT PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION GetTypeInfo (BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION GetIDsOfNames (BYVAL riid AS CONST IID CONST PTR, BYVAL rgszNames AS LPOLESTR PTR, BYVAL cNames AS UINT, BYVAL lcid AS LCID, BYVAL rgDispId AS DISPID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION Invoke (BYVAL dispIdMember AS DISPID, BYVAL riid AS CONST IID CONST PTR, BYVAL lcid AS LCID, BYVAL wFlags AS WORD, BYVAL pDispParams AS DISPPARAMS PTR, BYVAL pVarResult AS VARIANT PTR, BYVAL pExcepInfo AS EXCEPINFO PTR, BYVAL puArgErr AS UINT PTR) AS HRESULT
   DECLARE CONSTRUCTOR (BYVAL hwndContainer AS HWND)
   DECLARE DESTRUCTOR

   ' // Handle of the container's window
   m_hwndContainer AS HWND                                  ' // Handle of the OLE Container window
   m_pWebBrowser AS Afx_IWebBrowser2 PTR                    ' // Pointer to the IWebBrowser interface
   m_pHTMLDocument2 AS IHTMLDocument2 PTR                   ' // Pointer to the IHTMLDocument interface
   m_pHTMLDocumentEvents2 AS CHTMLDocumentEvents2Impl PTR   ' // Pointer to the CHTMLDocumentEvents2Impl class
   m_dwDocEventsCookie AS DWORD                             ' // Events cookie

   ' // Pointers to user implemented event procedures
   m_StatusTextChangeProc AS SUB (BYVAL hwndContainer AS HWND, BYVAL pwszText AS WSTRING PTR)
   m_DownloadCompleteProc AS SUB (BYVAL hwndContainer AS HWND)
   m_CommandStateChangeProc AS SUB (BYVAL hwndContainer AS HWND, BYVAL nCommand AS LONG, BYVAL fEnable AS VARIANT_BOOL)
   m_DownloadBeginProc AS SUB (BYVAL hwndContainer AS HWND)
   m_ProgressChangeProc AS SUB (BYVAL hwndContainer AS HWND, BYVAL Progress AS LONG, BYVAL ProgressMax AS LONG)
   m_PropertyChangeProc AS SUB (BYVAL hwndContainer AS HWND, BYVAL pwszProperty AS WSTRING PTR)
   m_TitleChangeProc AS SUB (BYVAL hwndContainer AS HWND, BYVAL pwszText AS WSTRING PTR)
   m_PrintTemplateInstantiationProc AS SUB (BYVAL hwndContainer AS HWND, BYVAL pdisp AS IDispatch PTR)
   m_PrintTemplateTeardownProc AS SUB (BYVAL hwndContainer AS HWND, BYVAL pdisp AS IDispatch PTR)
   m_BeforeNavigate2Proc AS SUB (BYVAL hwndContainer AS HWND, BYVAL pdisp AS IDispatch PTR, _
      BYVAL vUrl AS VARIANT PTR, BYVAL Flags AS VARIANT PTR, BYVAL TargetFrameName AS VARIANT PTR, _
      BYVAL PostData AS VARIANT PTR, BYVAL Headers AS VARIANT PTR, BYVAL pbCancel AS VARIANT_BOOL PTR)
   m_NavigateComplete2Proc AS SUB (BYVAL hwndContainer AS HWND, BYVAL pdisp AS IDispatch PTR, BYVAL vUrl AS VARIANT PTR)
   m_OnVisibleProc AS SUB (BYVAL hwndContainer AS HWND, BYVAL bVisible AS VARIANT_BOOL)
   m_DocumentCompleteProc AS SUB (BYVAL hwndContainer AS HWND, BYVAL pdisp AS IDispatch PTR, BYVAL vUrl AS VARIANT PTR)
   m_WindowSetResizableProc AS SUB (BYVAL hwndContainer AS HWND, BYVAL bResizable AS VARIANT_BOOL)
   m_WindowClosingProc AS SUB (BYVAL hwndContainer AS HWND, BYVAL IsChildWindow AS VARIANT_BOOL, BYVAL pbCancel AS VARIANT_BOOL PTR)
   m_WindowSetLeftProc AS SUB (BYVAL hwndContainer AS HWND, BYVAL nLenft AS LONG)
   m_WindowSetTopProc AS SUB (BYVAL hwndContainer AS HWND, BYVAL nTop AS LONG)
   m_WindowSetWidthProc AS SUB (BYVAL hwndContainer AS HWND, BYVAL nWidth AS LONG)
   m_WindowSetHeightProc AS SUB (BYVAL hwndContainer AS HWND, BYVAL nHeight AS LONG)
   m_ClientToHostWindowProc AS SUB (BYVAL hwndContainer AS HWND, BYVAL CX AS LONG PTR, BYVAL CY AS LONG PTR)
   m_SetSecureLockIconProc AS SUB (BYVAL hwndContainer AS HWND, BYVAL SecureLockIcon AS LONG)
   m_FileDownloadProc AS SUB (BYVAL hwndContainer AS HWND, BYVAL pbCancel AS VARIANT_BOOL PTR)
   m_NavigateErrorProc AS SUB (BYVAL hwndContainer AS HWND, BYVAL pdisp AS IDispatch PTR, _
      BYVAL vUrl AS VARIANT PTR, BYVAL vFrame AS VARIANT PTR, BYVAL StatusCode AS VARIANT PTR, BYVAL pbCancel AS VARIANT_BOOL PTR)
   m_PrivacyImpactedStateChangeProc AS SUB (BYVAL hwndContainer AS HWND, BYVAL bImpacted AS VARIANT_BOOL)
   m_NewWindow3Proc AS SUB (BYVAL hwndContainer AS HWND, BYVAL ppdispVal AS IDispatch PTR PTR, _
      BYVAL pbCancel AS VARIANT_BOOL PTR, BYVAL dwFlags AS ULONG, BYVAL pwszUrlContext AS WSTRING PTR, BYVAL pwszUrl AS WSTRING PTR)
   m_WindowStateChangedProc AS SUB (BYVAL hwndContainer AS HWND, BYVAL dwFlags AS LONG, BYVAL dwValidFlagsMask AS LONG)

END TYPE
' ########################################################################################

' ========================================================================================
' Constructors
' ========================================================================================
CONSTRUCTOR CWebBrowserEventsImpl (BYVAL hwndContainer AS HWND)
   CWB_DP("CONSTRUCTOR CWebBrowserEventsImpl")
   ' // Create an instance of the CHTMLDocumentEvents2Impl class
   m_hwndContainer = hwndContainer
   m_pHTMLDocumentEvents2 = NEW CHTMLDocumentEvents2Impl(m_hwndContainer, m_pWebBrowser)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destructor
' ========================================================================================
DESTRUCTOR CWebBrowserEventsImpl
   CWB_DP("DESTRUCTOR CWebBrowserEventsImpl")
   ' // Delete the HTMLDocumentEvents2 class
   IF m_pHTMLDocument2 THEN m_pHTMLDocument2->lpvtbl->Release(m_pHTMLDocument2)
   IF m_pHTMLDocumentEvents2 THEN Delete m_pHTMLDocumentEvents2
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns a pointer to our implementation of the WebBrowser events class.
' ========================================================================================
FUNCTION CWebBrowserEventsImpl.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   CWB_DP("CWebBrowserEventsImpl.QueryInterface")
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   IF IsEqualIID(riid, @DIID_DWebBrowserEvents2) OR _
      IsEqualIID(riid, @IID_IUnknown) OR _
      IsEqualIID(riid, @IID_IDispatch) THEN
      *ppvObj = @this
      ' // Not really needed, since this is not a COM object
      cast(Afx_IUnknown PTR, *ppvObj)->AddRef
      RETURN S_OK
   END IF
   RETURN E_NOINTERFACE
END FUNCTION
' ========================================================================================

' ========================================================================================
' This is not a COM class, so the life of the object is not managed by a reference count.
' ========================================================================================
FUNCTION CWebBrowserEventsImpl.AddRef () AS ULONG
   CWB_DP("CWebBrowserEventsImpl.AddRef")
   RETURN 1
END FUNCTION
' ========================================================================================

' ========================================================================================
' This is not a COM class, so the life of the object is not managed by a reference count.
' ========================================================================================
FUNCTION CWebBrowserEventsImpl.Release () AS ULONG
   CWB_DP("CWebBrowserEventsImpl.Release")
   RETURN 1
END FUNCTION
' ========================================================================================

' ========================================================================================
' The number of type information interfaces provided by the object. If the object provides
' type information, this number is 1; otherwise the number is 0.
' ========================================================================================
FUNCTION CWebBrowserEventsImpl.GetTypeInfoCount (BYVAL pctInfo AS UINT PTR) AS HRESULT
   CWB_DP("CWebBrowserEventsImpl.GetTypeInfoCount")
   *pctInfo = 0
   RETURN E_NOTIMPL
END FUNCTION
' ========================================================================================

' ========================================================================================
' The type information for an object, which can then be used to get the type information
' for an interface.
' ========================================================================================
FUNCTION CWebBrowserEventsImpl.GetTypeInfo (BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   CWB_DP("CWebBrowserEventsImpl.GetTypeInfo")
   RETURN E_NOTIMPL
END FUNCTION
' ========================================================================================

' ========================================================================================
' Maps a single member and an optional set of argument names to a corresponding set of
' integer DISPIDs, which can be used on subsequent calls to Invoke.
' ========================================================================================
FUNCTION CWebBrowserEventsImpl.GetIDsOfNames (BYVAL riid AS CONST IID CONST PTR, BYVAL rgszNames AS LPOLESTR PTR, BYVAL cNames AS UINT, BYVAL lcid AS LCID, BYVAL rgDispId AS DISPID PTR) AS HRESULT
   CWB_DP("CWebBrowserEventsImpl.GetIDsOfNames")
   RETURN E_NOTIMPL
END FUNCTION
' ========================================================================================

' ========================================================================================
' Provides access to properties and methods exposed by an object.
' Parameters in the rgvarg array of variants of the DISPPARAMS structures are zero based
' and in reverse order.
' ========================================================================================
FUNCTION CWebBrowserEventsImpl.Invoke (BYVAL dispIdMember AS DISPID, BYVAL riid AS CONST IID CONST PTR, BYVAL lcid AS LCID, BYVAL wFlags AS WORD, BYVAL pDispParams AS DISPPARAMS PTR, BYVAL pVarResult AS VARIANT PTR, BYVAL pExcepInfo AS EXCEPINFO PTR, BYVAL puArgErr AS UINT PTR) AS HRESULT
   CWB_DP("CWebBrowserEventsImpl.Invoke dispid = " & WSTR(dispIdMember))

   SELECT CASE dispIdMember

      ' ----------------------------------------------------------------------------------
      CASE 102   ' // StatusTextChange
      ' ----------------------------------------------------------------------------------
      CWB_DP("CWebBrowserEventsImpl.StatusTextChange")
      ' // Fires when the status bar text of the object has changed.
      ' // You can use the information provided by this event to update the text of a status bar.
      ' // Remarks: Do not free the BSTR (it is owned by the caller).
      IF m_StatusTextChangeProc THEN m_StatusTextChangeProc(m_hwndContainer, pDispParams->rgvarg[0].bstrVal)

      ' ----------------------------------------------------------------------------------
      CASE 104   ' // DownloadComplete
      ' ----------------------------------------------------------------------------------
      CWB_DP("CWebBrowserEventsImpl.DownloadComplete")
      ' // Fires when a navigation operation finishes, is halted, or fails.
      IF m_DownloadCompleteProc THEN m_DownloadCompleteProc(m_hwndContainer)

      ' ----------------------------------------------------------------------------------
      CASE 105   ' // CommandStateChange
      ' ----------------------------------------------------------------------------------
      CWB_DP("CWebBrowserEventsImpl.CommandStateChange")
      ' // Fires when the enabled state of a command changes.
      IF m_CommandStateChangeProc THEN m_CommandStateChangeProc(m_hwndContainer, pDispParams->rgvarg[1].lVal, pDispParams->rgvarg[0].boolVal)

      ' ----------------------------------------------------------------------------------
      CASE 106   ' // DownloadBegin
      ' ----------------------------------------------------------------------------------
      CWB_DP("CWebBrowserEventsImpl.DownloadBegin")
      ' // Fires when a navigation operation is beginning.
      IF m_DownloadBeginProc THEN m_DownloadBeginProc(m_hwndContainer)

      ' ----------------------------------------------------------------------------------
      CASE 108   ' // ProgressChange
      ' ----------------------------------------------------------------------------------
      CWB_DP("CWebBrowserEventsImpl.ProgressChange")
      ' // Fires when the progress of a download operation is updated on the object.
      ' // The container can use the information provided by this event to display the number of bytes
      ' // downloaded so far or to update a progress indicator.
      ' // To calculate the percentage of progress to show in a progress indicator, multiply the value
      ' // of Progress by 100 and divide by the value of ProgressMax (unless Progress is -1, in which
      ' // case the container can indicate that the operation is finished or hide the progress indicator).
      IF m_ProgressChangeProc THEN m_ProgressChangeProc(m_hwndContainer, pDispParams->rgvarg[1].lVal, pDispParams->rgvarg[0].lVal)

      ' ----------------------------------------------------------------------------------
      CASE 112   ' // PropertyChange
      ' ----------------------------------------------------------------------------------
      CWB_DP("CWebBrowserEventsImpl.PropertyChange")
      ' // Fires when the PutProperty method of the object changes the value of a property.
      IF m_PropertyChangeProc THEN m_PropertyChangeProc(m_hwndContainer, pDispParams->rgvarg[0].bstrVal)

      ' ----------------------------------------------------------------------------------
      CASE 113   ' // TitleChange
      ' ----------------------------------------------------------------------------------
      CWB_DP("CWebBrowserEventsImpl.TitleChange")
      ' // Fires when the title of a document in the object becomes available or changes.
      ' // Because the title might change while an HTML page is downloading, the URL of the
      ' // document is set as the title. After the title specified in the HTML page, if
      ' // there is one, is parsed, the title is changed to reflect the actual title.
      IF m_TitleChangeProc THEN m_TitleChangeProc(m_hwndContainer, pDispParams->rgvarg[0].bstrVal)

      ' ----------------------------------------------------------------------------------
      CASE 225   ' // PrintTemplateInstantiation
      ' ----------------------------------------------------------------------------------
      CWB_DP("CWebBrowserEventsImpl.PrintTemplateInstantiation")
      ' // Fires when a print template has been instantiated.
      IF m_PrintTemplateInstantiationProc THEN m_PrintTemplateInstantiationProc(m_hwndContainer, pDispParams->rgvarg[0].pdispVal)

      ' ----------------------------------------------------------------------------------
      CASE 226   ' // PrintTemplateTeardown
      ' ----------------------------------------------------------------------------------
      CWB_DP("CWebBrowserEventsImpl.PrintTemplateTeardown")
      ' // Fires when a print template has been destroyed.
      IF m_PrintTemplateTeardownProc THEN m_PrintTemplateTeardownProc(m_hwndContainer, pDispParams->rgvarg[0].pdispVal)

      ' ----------------------------------------------------------------------------------
      CASE 227   ' // UpdatePageStatus
      ' ----------------------------------------------------------------------------------
      CWB_DP("CWebBrowserEventsImpl.UpdatePageStatus")
      ' // Not implemented.

      ' ----------------------------------------------------------------------------------
      CASE 250   ' // BeforeNavigate2
      ' ----------------------------------------------------------------------------------
      CWB_DP("CWebBrowserEventsImpl.BeforeNavigate2")
      ' // Fires before navigation occurs in the given object (on either a window or frameset element).
      IF m_BeforeNavigate2Proc THEN m_BeforeNavigate2Proc(m_hwndContainer, pDispParams->rgvarg[6].pdispVal,_
         @pDispParams->rgvarg[5], @pDispParams->rgvarg[4], @pDispParams->rgvarg[3], _
         @pDispParams->rgvarg[2], @pDispParams->rgvarg[1], @pDispParams->rgvarg[0].boolVal)

      ' ----------------------------------------------------------------------------------
      CASE 252   ' // NavigateComplete2
      ' ----------------------------------------------------------------------------------
      CWB_DP("CWebBrowserEventsImpl.NavigateComplete2")
      ' // Fires after a navigation to a link is completed on either a window or frameset element.
      IF m_NavigateComplete2Proc THEN m_NavigateComplete2Proc(m_hwndContainer, pDispParams->rgvarg[1].pdispVal, @pDispParams->rgvarg[0])

      ' ----------------------------------------------------------------------------------
      CASE 253   ' // OnQuit
      ' ----------------------------------------------------------------------------------
      CWB_DP("CWebBrowserEventsImpl.OnQuit")
      ' // Fires before the Windows Internet Explorer application quits.
      ' // The WebBrowser object ignores this event.

      ' ----------------------------------------------------------------------------------
      CASE 254   ' // OnVisible
      ' ----------------------------------------------------------------------------------
      CWB_DP("CWebBrowserEventsImpl.OnVisible")
      ' //  Fires when the IWebBrowser2.Visible property of the object is changed.
      IF m_OnVisibleProc THEN m_OnVisibleProc(m_hwndContainer, pDispParams->rgvarg[0].boolVal)

      ' ----------------------------------------------------------------------------------
      CASE 259   ' // DocumentComplete
      ' ----------------------------------------------------------------------------------
      CWB_DP("CWebBrowserEventsImpl.DocumentComplete")
      ' // Fires when a document is completely loaded and initialized.
      IF m_DocumentCompleteProc THEN m_DocumentCompleteProc(m_hwndContainer, pDispParams->rgvarg[1].pdispVal, @pDispParams->rgvarg[0])

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
         pWebBrowser2->get_Document(@cast(IDispatch PTR, m_pHTMLDocument2))
         IF m_pHTMLDocument2 = NULL THEN EXIT FUNCTION
         AfxAdvise(m_pHTMLDocument2, m_pHTMLDocumentEvents2, DIID_HTMLDocumentEvents2_, @m_dwDocEventsCookie)
      END IF
      ' // ---------------------------------------------------------s----------------------

      ' ----------------------------------------------------------------------------------
      CASE 262   ' // WindowSetResizable
      ' ----------------------------------------------------------------------------------
      CWB_DP("CWebBrowserEventsImpl.WindowSetResizable")
      ' // Fires to indicate whether the host window should allow or disallow resizing of the object.
      IF m_WindowSetResizableProc THEN m_WindowSetResizableProc(m_hwndContainer, pDispParams->rgvarg[0].boolVal)

      ' ----------------------------------------------------------------------------------
      CASE 263   ' // WindowClosing
      ' ----------------------------------------------------------------------------------
      CWB_DP("CWebBrowserEventsImpl.WindowClosing")
      ' // Fires when the window of the object is about to be closed by script.
      IF m_WindowClosingProc THEN m_WindowClosingProc(m_hwndContainer, pDispParams->rgvarg[1].boolVal, @pDispParams->rgvarg[0].boolVal)

      ' ----------------------------------------------------------------------------------
      CASE 264   ' // WindowSetLeft
      ' ----------------------------------------------------------------------------------
      CWB_DP("CWebBrowserEventsImpl.WindowSetLeft")
      ' // Fires when the object changes its left position.
      IF m_WindowSetLeftProc THEN m_WindowSetLeftProc(m_hwndContainer, pDispParams->rgvarg[0].lVal)

      ' ----------------------------------------------------------------------------------
      CASE 265   ' // WindowSetTop
      ' ----------------------------------------------------------------------------------
      CWB_DP("CWebBrowserEventsImpl.WindowSetTop")
      ' // Fires when the object changes its top position.
      IF m_WindowSetTopProc THEN m_WindowSetTopProc(m_hwndContainer, pDispParams->rgvarg[0].lVal)

      ' ----------------------------------------------------------------------------------
      CASE 266   ' // WindowSetWidth
      ' ----------------------------------------------------------------------------------
      CWB_DP("CWebBrowserEventsImpl.WindowSetWidth")
      ' // Fires when the object changes its width.
      IF m_WindowSetWidthProc THEN m_WindowSetWidthProc(m_hwndContainer, pDispParams->rgvarg[0].lVal)

      ' ----------------------------------------------------------------------------------
      CASE 267   ' // WindowSetHeight
      ' ----------------------------------------------------------------------------------
      CWB_DP("CWebBrowserEventsImpl.WindowSetHeight")
      ' // Fires when the object changes its height.
      IF m_WindowSetHeightProc THEN m_WindowSetHeightProc(m_hwndContainer, pDispParams->rgvarg[0].lVal)

      ' ----------------------------------------------------------------------------------
      CASE 268   ' // ClientToHostWindow
      ' ----------------------------------------------------------------------------------
      CWB_DP("CWebBrowserEventsImpl.ClientToHostWindow")
      ' // Fires to request that the client window size be converted to the host window size.
      IF m_ClientToHostWindowProc THEN m_ClientToHostWindowProc(m_hwndContainer, @pDispParams->rgvarg[1].lVal, @pDispParams->rgvarg[0].lVal)

      ' ----------------------------------------------------------------------------------
      CASE 269   ' // SetSecureLockIcon
      ' ----------------------------------------------------------------------------------
      CWB_DP("CWebBrowserEventsImpl.SetSecureLockIcon")
      ' // Fires when there is a change in encryption level.
      IF m_SetSecureLockIconProc THEN m_SetSecureLockIconProc(m_hwndContainer, pDispParams->rgvarg[0].lVal)

      ' ----------------------------------------------------------------------------------
      CASE 270   ' // FileDownload
      ' ----------------------------------------------------------------------------------
      CWB_DP("CWebBrowserEventsImpl.FileDownload")
      ' // Fires to indicate that a file download is about to occur. If a file download
      ' // dialog is to be displayed, this event is fired prior to the display of the dialog.
      IF m_FileDownloadProc THEN m_FileDownloadProc(m_hwndContainer, @pDispParams->rgvarg[0].boolVal)

      ' ----------------------------------------------------------------------------------
      CASE 271   ' // NavigateError
      ' ----------------------------------------------------------------------------------
      CWB_DP("CWebBrowserEventsImpl.NavigateError")
      ' // Fires when an error occurs during navigation.
      IF m_NavigateErrorProc THEN m_NavigateErrorProc(m_hwndContainer, pDispParams->rgvarg[4].pdispVal, _
         @pDispParams->rgvarg[3], @pDispParams->rgvarg[2], @pDispParams->rgvarg[1], @pDispParams->rgvarg[0].boolVal)

      ' ----------------------------------------------------------------------------------
      CASE 272   ' // PrivacyImpactedStateChange
      ' ----------------------------------------------------------------------------------
      CWB_DP("CWebBrowserEventsImpl.PrivacyImpactedStateChange")
      ' // Fired when an event occurs that impacts privacy or when a user navigates away
      ' // from a URL that has impacted privacy.
      ' DIM bImpacted AS BOOLEAN = pDispParams->rgvarg[0].boolVal
      IF m_PrivacyImpactedStateChangeProc THEN m_PrivacyImpactedStateChangeProc(m_hwndContainer, pDispParams->rgvarg[0].boolVal)

      ' ----------------------------------------------------------------------------------
      CASE 273   ' // NewWindow3
      ' ----------------------------------------------------------------------------------
      CWB_DP("CWebBrowserEventsImpl.NewWindow3")
      ' // Raised when a new window is to be created.
      IF m_NewWindow3Proc THEN m_NewWindow3Proc(m_hwndContainer, @pDispParams->rgvarg[4].pdispVal, _
         @pDispParams->rgvarg[3].boolVal, pDispParams->rgvarg[2].ulVal, pDispParams->rgvarg[1].bstrVal, _
         pDispParams->rgvarg[0].bstrVal)

      ' ----------------------------------------------------------------------------------
      CASE 283   ' // WindowStateChanged
      ' ----------------------------------------------------------------------------------
      CWB_DP("CWebBrowserEventsImpl.WindowStateChanged")
      ' // Fires when the visibility state of a content window, such as the browser window
      ' // or a tab, changes.
      IF m_WindowStateChangedProc THEN m_WindowStateChangedProc(m_hwndContainer, pDispParams->rgvarg[1].lVal, pDispParams->rgvarg[0].lVal)

   END SELECT

   RETURN S_OK

END FUNCTION
' ========================================================================================
