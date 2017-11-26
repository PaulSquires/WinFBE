' ########################################################################################
' CHTMLDocumentEvents2 class
' Implementation of the HTMLDocumentEvents2 events sink class
' ########################################################################################
TYPE CHTMLDocumentEvents2 EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULONG
   DECLARE VIRTUAL FUNCTION GetTypeInfoCount (BYVAL pctinfo AS UINT PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION GetTypeInfo (BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION GetIDsOfNames (BYVAL riid AS CONST IID CONST PTR, BYVAL rgszNames AS LPOLESTR PTR, BYVAL cNames AS UINT, BYVAL lcid AS LCID, BYVAL rgDispId AS DISPID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION Invoke (BYVAL dispIdMember AS DISPID, BYVAL riid AS CONST IID CONST PTR, BYVAL lcid AS LCID, BYVAL wFlags AS WORD, BYVAL pDispParams AS DISPPARAMS PTR, BYVAL pVarResult AS VARIANT PTR, BYVAL pExcepInfo AS EXCEPINFO PTR, BYVAL puArgErr AS UINT PTR) AS HRESULT

   DECLARE CONSTRUCTOR (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pWebBrowser AS Afx_IWebBrowser2 PTR)
   DECLARE DESTRUCTOR

   m_pWebCtx AS CWebCtx PTR                ' // Pointer to the OLE container class
   m_pWebBrowser AS Afx_IWebBrowser2 PTR   ' // Pointer to the IWebBrowser interface
   cRef AS ULONG                           ' // Reference count
   m_pHtmlDocumentEventsProc AS FUNCTION (BYVAL pWebCtx AS CWebCtx PTR, BYVAL dispId AS LONG, BYVAL pEvtObj AS IHTMLEventObj PTR) AS BOOLEAN

END TYPE
' ########################################################################################

' ========================================================================================
' Constructors
' ========================================================================================
CONSTRUCTOR CHTMLDocumentEvents2 (BYVAL pWebCtx AS CWebCtx PTR, BYVAL pWebBrowser AS Afx_IWebBrowser2 PTR)
   CWBX_DP("CONSTRUCTOR CHTMLDocumentEvents2")
   m_pWebCtx = pWebCtx
   m_pWebBrowser = pWebBrowser
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destructor
' ========================================================================================
DESTRUCTOR CHTMLDocumentEvents2
   CWBX_DP("DESTRUCTOR CHTMLDocumentEvents2")
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns a pointer to our implementation of the WebBrowser events class.
' ========================================================================================
FUNCTION CHTMLDocumentEvents2.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   CWBX_DP("CHTMLDocumentEvents2.QueryInterface")
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   ' extern DIID_HTMLDocumentEvents2 is broken and causes the linker to fail
   DIM DIID_HTMLDocumentEvents2_ AS GUID = (&h3050F613, &h98B5, &h11CF, {&hBB, &h82, &h00, &hAA, &h00, &hBD, &hCE, &h0B})
   IF IsEqualIID(riid, @DIID_HTMLDocumentEvents2_) OR _
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
FUNCTION CHTMLDocumentEvents2.AddRef () AS ULONG
   this.cRef += 1
   CWBX_DP("CHTMLDocumentEvents2 AddRef = " & WSTR(this.cRef))
   RETURN this.cRef
END FUNCTION
' ========================================================================================

' ========================================================================================
' This is not a COM class, so the life of the object is not managed by a reference count.
' ========================================================================================
FUNCTION CHTMLDocumentEvents2.Release () AS ULONG
   this.cRef -= 1
   CWBX_DP("CHTMLDocumentEvents2 Release = " & WSTR(this.cRef))
   FUNCTION = this.cRef
   ' // Don't delete it because this is not a COM class
   'IF this.cRef = 0 THEN DELETE @this
END FUNCTION
' ========================================================================================

' ========================================================================================
' The number of type information interfaces provided by the object. If the object provides
' type information, this number is 1; otherwise the number is 0.
' ========================================================================================
FUNCTION CHTMLDocumentEvents2.GetTypeInfoCount (BYVAL pctInfo AS UINT PTR) AS HRESULT
   CWBX_DP("CHTMLDocumentEvents2.GetTypeInfoCount")
   *pctInfo = 0
   RETURN E_NOTIMPL
END FUNCTION
' ========================================================================================

' ========================================================================================
' The type information for an object, which can then be used to get the type information
' for an interface.
' ========================================================================================
FUNCTION CHTMLDocumentEvents2.GetTypeInfo (BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   CWBX_DP("CHTMLDocumentEvents2.GetTypeInfo")
   RETURN E_NOTIMPL
END FUNCTION
' ========================================================================================

' ========================================================================================
' Maps a single member and an optional set of argument names to a corresponding set of
' integer DISPIDs, which can be used on subsequent calls to Invoke.
' ========================================================================================
FUNCTION CHTMLDocumentEvents2.GetIDsOfNames (BYVAL riid AS CONST IID CONST PTR, BYVAL rgszNames AS LPOLESTR PTR, BYVAL cNames AS UINT, BYVAL lcid AS LCID, BYVAL rgDispId AS DISPID PTR) AS HRESULT
   CWBX_DP("CHTMLDocumentEvents2.GetIDsOfNames")
   RETURN E_NOTIMPL
END FUNCTION
' ========================================================================================

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

' ========================================================================================
' Provides access to properties and methods exposed by an object.
' Parameters in the rgvarg array of variants of the DISPPARAMS structures are zero based
' and in reverse order.
' Event handlers for cancelable document events must return a Boolean value in the pvarResult
' parameter of the IDispatch::Invoke() method. A value of VARIANT_TRUE indicates that
' Internet Explorer should perform its own event processing; a value of VARIANT_FALSE
' cancels the event.
' ========================================================================================
FUNCTION CHTMLDocumentEvents2.Invoke (BYVAL dispIdMember AS DISPID, BYVAL riid AS CONST IID CONST PTR, BYVAL lcid AS LCID, BYVAL wFlags AS WORD, BYVAL pDispParams AS DISPPARAMS PTR, BYVAL pVarResult AS VARIANT PTR, BYVAL pExcepInfo AS EXCEPINFO PTR, BYVAL puArgErr AS UINT PTR) AS HRESULT
   CWBX_DP("CHTMLDocumentEvents2.Invoke dispid = " & WSTR(dispIdMember))

   ' // Send the event for processing
'   IF m_pHtmlDocumentEventsProc THEN m_pHtmlDocumentEventsProc(m_pAxHost, dispIdMember, CAST(IHTMLEventObj PTR, pDispParams->rgvarg[0].pdispVal))

   ' Event handlers for cancelable document events must return a Boolean value in the pvarResult
   ' parameter of the IDispatch.Invoke() method. A value of VARIANT_TRUE indicates that
   ' Internet Explorer should perform its own event processing; a value of VARIANT_FALSE
   ' cancels the event.
   ' See: https://support.microsoft.com/en-us/kb/246247
   IF m_pHtmlDocumentEventsProc THEN
      DIM bEventRet AS BOOLEAN
      bEventRet = m_pHtmlDocumentEventsProc(m_pWebCtx, dispIdMember, CAST(IHTMLEventObj PTR, pDispParams->rgvarg[0].pdispVal))
      IF bEventRet THEN
         pvarResult->vt = VT_BOOL
         pvarResult->boolVal = bEventRet
      END IF
   END IF

'   SELECT CASE dispIdMember

'      ' ----------------------------------------------------------------------------------
'      CASE -600   ' // onclick
'      ' ----------------------------------------------------------------------------------
'      CWBX_DP("CHTMLDocumentEvents2.Invoke - onclick")

   ' // Example to get the identifier of the clicked element, if it as one.
   ' // You can call other methods, like get_outerHtml.
   ' // See: https://msdn.microsoft.com/en-us/library/aa703876(v=vs.85).aspx
   '      ' // Get a reference to the element that has fired the event
   '      DIM pEvtObj AS IHTMLEventObj PTR = CAST(IHTMLEventObj PTR, pDispParams->rgvarg[0].pdispVal)
   '      DIM pElement AS IHTMLElement PTR
   '      IF pEvtObj THEN pEvtObj->lpvtbl->get_srcElement(pEvtObj, @pElement)
   '      IF pElement = NULL THEN EXIT FUNCTION
   '      DIM cbsId AS CBSTR
   '      pElement->lpvtbl->get_id(pElement, @cbsID)
   '      pElement->lpvtbl->Release(pElement)
   '      AfxMsg cbsID

'   END SELECT

   RETURN S_OK

END FUNCTION
' ========================================================================================
