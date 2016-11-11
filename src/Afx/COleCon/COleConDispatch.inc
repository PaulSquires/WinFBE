' ########################################################################################
' Microsoft Windows
' File: COleConDispatch.inc
' Contents: OLE Container dispatch class
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once

' ========================================================================================
' Afx_OleConDispatch class
' ========================================================================================
TYPE Afx_OleConDispatch EXTENDS Afx_IDispatch

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULONG
   DECLARE VIRTUAL FUNCTION GetTypeInfoCount (BYVAL pctinfo AS UINT PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION GetTypeInfo (BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION GetIDsOfNames (BYVAL riid AS CONST IID CONST PTR, BYVAL rgszNames AS LPOLESTR PTR, BYVAL cNames AS UINT, BYVAL lcid AS LCID, BYVAL rgDispId AS DISPID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION Invoke (BYVAL dispIdMember AS DISPID, BYVAL riid AS CONST IID CONST PTR, BYVAL lcid AS LCID, BYVAL wFlags AS WORD, BYVAL pDispParams AS DISPPARAMS PTR, BYVAL pVarResult AS VARIANT PTR, BYVAL pExcepInfo AS EXCEPINFO PTR, BYVAL puArgErr AS UINT PTR) AS HRESULT

   m_hwnd                     AS HWND            ' // Window handle
   m_hFont                    AS HFONT           ' // Window font
   m_pOcx                     AS IDispatch PTR   ' // Address of a pointer to the ActiveX's Vtable
   m_pIFontDisp               AS IFontDisp PTR   ' // IFontDisp interface
   m_dwMiscStatus             AS DWORD           ' // Status bitwise flags
   m_pUnkSink                 AS IUnknown PTR    ' // IUnknown of the event class
   m_pEvtObj                  AS LPVOID          ' // Pointer to the event class
   m_bLocked                  AS INTEGER         ' // Container locked
   m_bInPlaceActive           AS INTEGER         ' // The control is in-place active
   m_bUIActive                AS INTEGER         ' // The control is UI Active
   m_bHaveFocus               AS INTEGER         ' // The control has focus
   m_AmbientDlControl         AS DWORD           ' // Ambient flags to control the WebBrowser behavior
   m_rxRatio                  AS SINGLE          ' // DPI horizontal ratio
   m_ryRatio                  AS SINGLE          ' // DPI vertical ratio
   m_wszProgID                AS WSTRING * 260   ' // ActiveX ProgID
   m_wszLibName               AS WSTRING * 260   ' // ActiveX path
   m_wszLicKey                AS WSTRING * 260   ' // ActiveX license key
   m_rclsid                   AS CLSID           ' // ActiveX ClsId
   m_riid                     AS IID             ' // ActiveX IID
   m_bIsRegFree               AS BOOLEAN         ' // Register free OCX
   m_bIsLicensed              AS BOOLEAN         ' // Licensed OCX

   ' *** Implemented interfaces ***
   m_pIOleClientSiteImpl      AS Afx_IOleClientSiteImpl_ PTR        ' // IOleClientSite interface (implemented)
   m_pIOleContainerImpl       AS Afx_IOleContainerImpl_ PTR         ' // IOleContainer interface (implemented)
   m_pIOleControlSiteImpl     AS Afx_IOleControlSiteImpl_ PTR       ' // IOleControlSite interface (implemented)
   m_pIOleInPlaceFrameImpl    AS Afx_IOleInPlaceFrameImpl_ PTR      ' // IOleInPlaceFrame interface (implemented)
   m_pIOleInPlaceSiteImpl     AS Afx_IOleInPlaceSiteImpl_ PTR       ' // IOleInPlaceSite interface (implemented)
   m_pIOleInplaceSiteExImpl   AS Afx_IOleInPlaceSiteExImpl_ PTR     ' // IOleInPlaceSiteEx interface (implemented)
   m_pIOleInPlaceUIWindowImpl AS Afx_IOleInplaceUIWindowImpl_ PTR   ' // IOleInplaceUIWindow interface (implemented)
   m_pIOleWindowImpl          AS Afx_IOleWindowImpl_ PTR            ' // IOleWindow interface (implemented)
   m_pIParseDisplayNameImpl   AS Afx_IParseDisplayNameImpl_ PTR     ' // IParseDisplayName interface (implemented)
   m_pIServiceProviderImpl    AS Afx_IServiceProviderImpl_ PTR      ' // IServiceProvider interface (implemented)
   m_pISimpleFrameSiteImpl    AS Afx_ISimpleFrameSiteImpl_ PTR      ' // ISimpleFrameSite interface (implemented)

   ' *** Control interfaces ***
   m_pIOleObject              AS Afx_IOleObject PTR                 ' // IOleObject interface
   m_pIOleInPlaceObject       AS Afx_IOleInPlaceObject PTR          ' // IOleInPlaceObject interface
   m_pIOleInPlaceActiveObject AS Afx_IOleInPlaceActiveObject PTR    ' // IOleInPlaceActiveObject interface

   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR
   DECLARE FUNCTION StrToGuid (BYVAL pwszGuidText AS WSTRING PTR = NULL) AS GUID
   DECLARE FUNCTION GuidText (BYVAL classID AS CLSID PTR) AS STRING
   DECLARE FUNCTION GuidText (BYVAL riid AS REFIID) AS STRING
   DECLARE FUNCTION GuidText (BYVAL classID AS CLSID) AS STRING
   DECLARE FUNCTION GuidFromStr (BYVAL pwszGuidText AS WSTRING PTR = NULL) AS GUID

   cRef AS ULONG

END TYPE
' ========================================================================================

#include once "Afx/COleCon/COleConSite.inc"
#include once "Afx/COleCon/CServiceProvider.inc"
#include once "Afx/COleCon/COleContainer.inc"
#include once "Afx/COleCon/COleInPlaceFrame.inc"
#include once "Afx/COleCon/COleInPlaceSite.inc"
#include once "Afx/COleCon/COleControlSite.inc"

' ========================================================================================
' Afx_OleConDispatch class
' Implements the IDispatch interface
' ========================================================================================
CONSTRUCTOR Afx_OleConDispatch
   OC_DP("Afx_OleConDispatch CONSTRUCTOR")
   m_pIOleClientSiteImpl      = NEW Afx_IOleClientSiteImpl(@this)
   m_pIOleContainerImpl       = NEW Afx_IOleContainerImpl(@this)
   m_pIOleControlSiteImpl     = NEW Afx_IOleControlSiteImpl(@this)
   m_pIOleInPlaceFrameImpl    = NEW Afx_IOleInPlaceFrameImpl(@this)
   m_pIOleInPlaceSiteImpl     = NEW Afx_IOleInPlaceSiteImpl(@this)
   m_pIOleInplaceSiteExImpl   = NEW Afx_IOleInPlaceSiteExImpl(@this)
   m_pIOleInPlaceUIWindowImpl = NEW Afx_IOleInplaceUIWindowImpl(@this)
   m_pIOleWindowImpl          = NEW Afx_IOleWindowImpl(@this)
   m_pIParseDisplayNameImpl   = NEW Afx_IParseDisplayNameImpl(@this)
   m_pIServiceProviderImpl    = NEW Afx_IServiceProviderImpl(@this)
   m_pISimpleFrameSiteImpl    = NEW Afx_ISimpleFrameSiteImpl(@this)
END CONSTRUCTOR

' ========================================================================================
' Cleanup
' ========================================================================================
DESTRUCTOR Afx_OleConDispatch
   OC_DP("Afx_OleConDispatch DESTRUCTOR")
   ' // Required: Deactivate the IOleInPlaceObject if it active.
   ' // It forces the release of the IOleInPlaceSite interface.
   IF m_bInPlaceActive THEN
      IF m_pIOleInPlaceObject THEN m_pIOleInPlaceObject->InPlaceDeactivate
   END IF
   ' // Close the control
   IF m_pIOleObject THEN
      m_pIOleObject->Close(OLECLOSE_NOSAVE)
      m_pIOleObject->SetClientSite(NULL)
   END IF
   DIM cRef AS DWORD
   ' // Release de IFontDisp object
   IF m_pIFontDisp THEN cRef = IUnknown_Release(m_pIFontDisp) : m_pIFontDisp = NULL
   OC_DP("Release IFontDisp - cRef (must be 0) = " & WSTR(cRef))
   ' // Release the IOleInPlaceObject interface
   IF m_pIOleInPlaceObject THEN cRef = m_pIOleInPlaceObject->Release : m_pIOleInPlaceObject = NULL
   OC_DP("Release IOleInPlaceObject - cRef (must be 0) = " & WSTR(cRef))
   ' // Release the IOleObject interface
   ' // The returned reference count must be at least 1 because the
   ' // ActiveX control holds a reference to it.
   IF m_pIOleObject THEN cRef = m_pIOleObject->Release : m_pIOleObject = NULL
   OC_DP("Release IOleObject - cRef (must be 1) = " & WSTR(cRef))
   ' // Release the ActiveX control
   IF m_pOcx THEN cRef = IUnknown_Release(m_pOcx) : m_pOcx = NULL
   OC_DP("Release OCX - cRef (must be 0) = " & WSTR(cRef))
   ' // Delete the implemented classess
   IF m_pIOleClientSiteImpl      THEN Delete m_pIOleClientSiteImpl       : m_pIOleClientSiteImpl = NULL
   IF m_pIOleContainerImpl       THEN Delete(m_pIOleContainerImpl)       : m_pIOleContainerImpl = NULL
   IF m_pIOleControlSiteImpl     THEN Delete(m_pIOleControlSiteImpl)     : m_pIOleControlSiteImpl = NULL
   IF m_pIOleInPlaceFrameImpl    THEN Delete(m_pIOleInPlaceFrameImpl)    : m_pIOleInPlaceFrameImpl = NULL
   IF m_pIOleInPlaceSiteImpl     THEN Delete(m_pIOleInPlaceSiteImpl)     : m_pIOleInPlaceSiteImpl = NULL
   IF m_pIOleInplaceSiteExImpl   THEN Delete(m_pIOleInPlaceSiteExImpl)   : m_pIOleInPlaceSiteExImpl = NULL
   IF m_pIOleInPlaceUIWindowImpl THEN Delete(m_pIOleInplaceUIWindowImpl) : m_pIOleInplaceUIWindowImpl = NULL
   IF m_pIOleWindowImpl          THEN Delete(m_pIOleWindowImpl)          : m_pIOleWindowImpl = NULL
   IF m_pIParseDisplayNameImpl   THEN Delete(m_pIParseDisplayNameImpl)   : m_pIParseDisplayNameImpl = NULL
   IF m_pIServiceProviderImpl    THEN Delete(m_pIServiceProviderImpl)    : m_pIServiceProviderImpl = NULL
   IF m_pISimpleFrameSiteImpl    THEN Delete(m_pISimpleFrameSiteImpl)    : m_pISimpleFrameSiteImpl = NULL
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
FUNCTION Afx_OleConDispatch.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   OC_DP("Afx_OleConDispatch QueryInterface : " & GuidText(riid))

   IF ppvObj = NULL THEN RETURN E_INVALIDARG

   IF IsEqualIID(riid, @IID_IUnknown) THEN
      *ppvObj = @this
   ELSEIF IsEqualIID(riid, @IID_IDispatch) THEN
      *ppvObj = @this
   ELSEIF IsEqualIID(riid, @IID_IOleClientSite) THEN
      *ppvObj = m_pIOleClientSiteImpl
   ELSEIF IsEqualIID(riid, @IID_IOleContainer) THEN
      *ppvObj = m_pIOleContainerImpl
   ELSEIF IsEqualIID(riid, @IID_IOleControlSite) THEN
      *ppvObj = m_pIOleControlSiteImpl
   ELSEIF IsEqualIID(riid, @IID_IOleInPlaceFrame) THEN
      *ppvObj = m_pIOleInPlaceFrameImpl
   ELSEIF IsEqualIID(riid, @IID_IOleInPlaceSite) THEN
      *ppvObj = m_pIOleInPlaceSiteImpl
   ELSEIF IsEqualIID(riid, @IID_IOleInPlaceSiteEx) THEN
      *ppvObj = m_pIOleInPlaceSiteExImpl
   ELSEIF IsEqualIID(riid, @IID_IOleInPlaceUIWindow) THEN
      *ppvObj = m_pIOleInPlaceUIWindowImpl
   ELSEIF IsEqualIID(riid, @IID_IOleWindow) THEN
      *ppvObj = m_pIOleWindowImpl
   ELSEIF IsEqualIID(riid, @IID_IParseDisplayName) THEN
      *ppvObj = m_pIParseDisplayNameImpl
   ELSEIF IsEqualIID(riid, @IID_IServiceProvider) THEN
      *ppvObj = m_pIServiceProviderImpl
   ELSEIF IsEqualIID(riid, @IID_ISimpleFrameSite) THEN
      *ppvObj = m_pISimpleFrameSiteImpl
   END IF

   IF *ppvObj = NULL THEN RETURN E_NOINTERFACE

   cast(Afx_IUnknown PTR, *ppvObj)->AddRef
   RETURN S_OK

END FUNCTION
' ========================================================================================

' ========================================================================================
' Increments the reference count for an interface on an object. This method should be called
' for every new copy of a pointer to an interface on an object.
' This is a dummy counter for debugging purposes since the lifetime of the OLE container
' classes, the host window and the hosted control is determined by the lifetime of the
' COleCon class. When the COleCon class is destroyed, the COM interfaces are released and
' the classes and host window destroyed.
' ========================================================================================
FUNCTION Afx_OleConDispatch.AddRef () AS ULONG
   cRef += 1
   RETURN cRef
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' ========================================================================================
FUNCTION Afx_OleConDispatch.Release () AS DWORD
   cRef -= 1
   RETURN cRef
END FUNCTION
' ========================================================================================

' ========================================================================================
' Not required. This method is necessary only for containers that support nonstandard
' ambient properties.
' ========================================================================================
FUNCTION Afx_OleConDispatch.GetTypeInfoCount (BYVAL pctInfo AS UINT PTR) AS HRESULT
   *pctInfo = 0
   RETURN E_NOTIMPL
END FUNCTION
' ========================================================================================

' ========================================================================================
' Not required. This method is necessary only for containers that support nonstandard
' ambient properties.
' ========================================================================================
FUNCTION Afx_OleConDispatch.GetTypeInfo (BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   RETURN E_NOTIMPL
END FUNCTION
' ========================================================================================

' ========================================================================================
' Not required. This method is necessary only for containers that support nonstandard
' ambient properties.
' ========================================================================================
FUNCTION Afx_OleConDispatch.GetIDsOfNames (BYVAL riid AS CONST IID CONST PTR, BYVAL rgszNames AS LPOLESTR PTR, BYVAL cNames AS UINT, BYVAL lcid AS LCID, BYVAL rgDispId AS DISPID PTR) AS HRESULT
   RETURN E_NOTIMPL
END FUNCTION
' ========================================================================================

' ========================================================================================
' Dispatches ambient properties
' ========================================================================================
' ========================================================================================
' Controlling Download and Execution (DISPID_AMBIENT_DLCONTROL)
' The WebBrowser Control gives you control over what it downloads, displays, and executes.
' To gain this control, you need to implement your host's IDispatch so it handles
' DISPID_AMBIENT_DLCONTROL. When the WebBrowser Control is instantiated, it will call your
' IDispatch.Invoke with this ID. Set pvarResult to a combination of following flags, using
' the bitwise OR operator, to indicate your preferences.
' - DLCTL_DLIMAGES, DLCTL_VIDEOS, and DLCTL_BGSOUNDS: Images, videos, and background sounds
'      will be downloaded from the server and displayed or played if these flags are set.
'      They will not be downloaded and displayed if the flags are not set.
' - DLCTL_NO_SCRIPTS and DLCTL_NO_JAVA: Scripts and Java applets will not be executed.
' - DLCTL_NO_DLACTIVEXCTLS and DLCTL_NO_RUNACTIVEXCTLS : ActiveX controls will not be
'      downloaded or will not be executed.
' - DLCTL_DOWNLOADONLY: The page will only be downloaded, not displayed.
' - DLCTL_NO_FRAMEDOWNLOAD: The WebBrowser Control will download and parse a frameset, but
'      not the individual frame objects within the frameset.
' - DLCTL_RESYNCHRONIZE and DLCTL_PRAGMA_NO_CACHE: These flags cause cache refreshes.
'      With DLCTL_RESYNCHRONIZE, the server will be asked for update status. Cached files
'      will be used if the server indicates that the cached information is up-to-date.
'      With DLCTL_PRAGMA_NO_CACHE, files will be re-downloaded from the server regardless
'      of the update status of the files.
' - DLCTL_NO_BEHAVIORS: Behaviors are not downloaded and are disabled in the document.
' - DLCTL_NO_METACHARSET_HTML: Character sets specified in meta elements are suppressed.
' - DLCTL_URL_ENCODING_DISABLE_UTF8 and DLCTL_URL_ENCODING_ENABLE_UTF8: These flags
'      function similarly to the DOCHOSTUIFLAG_URL_ENCODING_DISABLE_UTF8 and
'      DOCHOSTUIFLAG_URL_ENCODING_ENABLE_UTF8 flags used with IDocHostUIHandler.GetHostInfo.
'      The difference is that the DOCHOSTUIFLAG flags are checked only when the WebBrowser
'      Control is first instantiated. The download flags here for the ambient property
'      change are checked whenever the WebBrowser Control needs to perform a download.
' - DLCTL_NO_CLIENTPULL: No client pull operations will be performed.
' - DLCTL_SILENT: No user interface will be displayed during downloads.
' - DLCTL_FORCEOFFLINE: The WebBrowser Control always operates in offline mode.
' - DLCTL_OFFLINEIFNOTCONNECTED and DLCTL_OFFLINE: These flags are the same. The WebBrowser
'      Control will operate in offline mode if not connected to the Internet.
' DISPID_AMBIENT_DLCONTROL and the flag values are defined in mshtmdid.h.
' Initially, when the function call to IDispatch.Invoke starts, the VARIANT to which the
' parameter pvarResult points is of type VT_EMPTY. You must switch the type to VT_I4 for
' any settings to have an effect. You can place your flag values in the lVal member
' of the VARIANT.
' Most of the flag values have negative effects, that is, they prevent behavior that
' normally happens. For instance, scripts are normally executed by the WebBrowser Control
' if you don't customize its behavior. But if you set the DLCTL_NOSCRIPTS flag, no
' scripts will execute in that instance of the control. However, three flags—DLCTL_DLIMAGES,
' DLCTL_VIDEOS, and DLCTL_BGSOUNDS—work exactly opposite. If you set flags at all, you
' must set these three for the WebBrowser Control to behave in its default manner vis-a-vis
' images, videos and sounds.
' ========================================================================================
FUNCTION Afx_OleConDispatch.Invoke (BYVAL dispIdMember AS DISPID, BYVAL riid AS CONST IID CONST PTR, BYVAL lcid AS LCID, BYVAL wFlags AS WORD, BYVAL pDispParams AS DISPPARAMS PTR, BYVAL pVarResult AS VARIANT PTR, BYVAL pExcepInfo AS EXCEPINFO PTR, BYVAL puArgErr AS UINT PTR) AS HRESULT

   ' // riid must be a null guid
   IF IsEqualIID(riid, @IID_NULL) = FALSE THEN RETURN DISP_E_UNKNOWNINTERFACE

   ' // Anything but a property get is invalid
   IF (wFlags AND DISPATCH_PROPERTYGET) <> DISPATCH_PROPERTYGET THEN RETURN DISP_E_MEMBERNOTFOUND

   ' // pvarResult can't be null for get properties
   IF pvarResult = NULL THEN RETURN E_INVALIDARG

   ' // Default error code
   FUNCTION = DISP_E_MEMBERNOTFOUND

   SELECT CASE dispIdMember
   	CASE -701   ' DISPID_AMBIENT_BACKCOLOR
         pvarResult->vt = VT_UI4
         pvarResult->lVal = GetSysColor(COLOR_WINDOW)
         RETURN NOERROR
   	CASE -703   ' DISPID_AMBIENT_FONT
         IF m_pIFontDisp <> NULL THEN
            ' Incremenet the reference count
             IUnknown_AddRef(m_pIFontDisp)
            ' Return a reference to the ambient font
            pvarResult->vt = VT_DISPATCH
            pvarResult->pdispVal = cast(IDispatch PTR, m_pIFontDisp)
            RETURN NOERROR
         END IF
   	CASE DISPID_AMBIENT_FORECOLOR '-704
         pvarResult->vt = VT_UI4
         pvarResult->lVal = GetSysColor(COLOR_WINDOWTEXT)
         RETURN NOERROR
   	CASE  DISPID_AMBIENT_LOCALEID '-705
         pvarResult->vt = VT_I4
         pvarResult->lVal = LOCALE_USER_DEFAULT
         RETURN NOERROR
   	CASE  DISPID_AMBIENT_MESSAGEREFLECT '-706
         pvarResult->vt = VT_BOOL
         pvarResult->boolVal = 0
         RETURN NOERROR
   	CASE  DISPID_AMBIENT_USERMODE   '  -709
         pvarResult->vt = VT_BOOL
         pvarResult->boolVal = -1
         RETURN NOERROR
   	CASE  DISPID_AMBIENT_UIDEAD   ' -710
         pvarResult->vt = VT_BOOL
         pvarResult->boolVal = 0
         RETURN NOERROR
   	CASE  DISPID_AMBIENT_SHOWGRABHANDLES ' -711
         pvarResult->vt = VT_BOOL
         pvarResult->boolVal = 0
         RETURN NOERROR
   	CASE  DISPID_AMBIENT_SHOWHATCHING ' -712
         pvarResult->vt = VT_BOOL
         pvarResult->boolVal = 0
         RETURN NOERROR
   	CASE  DISPID_AMBIENT_DISPLAYASDEFAULT  ' -713
         pvarResult->vt = VT_BOOL
         pvarResult->boolVal = 0
         RETURN NOERROR
   	CASE  DISPID_AMBIENT_SUPPORTSMNEMONICS  ' -714
         pvarResult->vt = VT_BOOL
         pvarResult->boolVal = -1
         RETURN NOERROR
   	CASE -5501 ' DISPID_AMBIENT_OFFLINEIFNOTCONNECTED
         pvarResult->vt = VT_BOOL
         pvarResult->boolVal = -1
         RETURN NOERROR
   	CASE -5502   '  DISPID_AMBIENT_SILENT
         pvarResult->vt = VT_BOOL
         pvarResult->boolVal = -1
         RETURN NOERROR
   	CASE -5512   ' DISPID_AMBIENT_DLCONTROL
         IF m_AmbientDlControl THEN
            pvarResult->vt = VT_I4
            pvarResult->lVal = m_AmbientDlControl
            RETURN NOERROR
         END IF
'      CASE -702   ' DISPID_AMBIENT_DISPLAYNAME
'      CASE -707   ' DISPID_AMBIENT_SCALEUNITS
'      CASE -708   ' DISPID_AMBIENT_TEXTALIGN
'      CASE -715   ' DISPID_AMBIENT_AUTOCLIP
'      CASE -716   ' DISPID_AMBIENT_APPEARANCE
'      CASE -725   ' DISPID_AMBIENT_CODEPAGE
'      CASE -726   ' DISPID_AMBIENT_PALETTE
'      CASE -727   ' DISPID_AMBIENT_CHARSET
'      CASE -732   ' DISPID_AMBIENT_RIGHTTOLEFT
'      CASE -733   ' DISPID_AMBIENT_TOPTOBOTTOM
   END SELECT

END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts a string into a 16-byte (128-bit) Globally Unique Identifier (GUID)
' To be valid, the string must contain exactly 32 hexadecimal digits, delimited by hyphens
' and enclosed by curly braces. For example: {B09DE715-87C1-11D1-8BE3-0000F8754DA1}
' If pwszGuidText is omited, OC_Guid generates a new unique guid.
' Remarks: I have need to call the UuidCreate function dynamically because, at the time of
' writing, the library for the RPCRT4.DLL seems broken and the linker fails.
' ========================================================================================
PRIVATE FUNCTION Afx_OleConDispatch.StrToGuid (BYVAL pwszGuidText AS WSTRING PTR = NULL) AS GUID
   DIM rguid AS GUID
   IF pwszGuidText = NULL THEN
      ' // Generate a new guid
      DIM AS ANY PTR pLib = DyLibLoad("RPCRT4.DLL")
      IF pLib  THEN
         DIM pProc AS FUNCTION (BYVAL Uuid AS UUID PTR) AS RPC_STATUS
         pProc = DyLibSymbol(pLib, "UuidCreate")
         IF pProc THEN pProc(@rguid)
         DyLibFree(pLib)
      END IF
   ELSE
      CLSIDFromString(pwszGuidText, @rGuid)
   END IF
   FUNCTION = rguid
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a 38-byte human-readable guid string from a 16-byte GUID.
' ========================================================================================
PRIVATE FUNCTION Afx_OleConDispatch.GuidText (BYVAL classID AS CLSID PTR) AS STRING
   DIM pwsz AS WSTRING PTR
   StringFromCLSID(classID, CAST(LPOLESTR PTR, @pwsz))
   FUNCTION = *pwsz
   CoTaskMemFree(pwsz)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Afx_OleConDispatch.GuidText (BYVAL riid AS REFIID) AS STRING
   DIM pwsz AS WSTRING PTR
   StringFromIID(riid, CAST(LPOLESTR PTR, @pwsz))
   FUNCTION = *pwsz
   CoTaskMemFree(pwsz)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION Afx_OleConDispatch.GuidText (BYVAL classID AS CLSID) AS STRING
   DIM pwsz AS WSTRING PTR
   StringFromCLSID(@classID, CAST(LPOLESTR PTR, @pwsz))
   FUNCTION = *pwsz
   CoTaskMemFree(pwsz)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts a string into a 16-byte (128-bit) Globally Unique Identifier (GUID)
' To be valid, the string must contain exactly 32 hexadecimal digits, delimited by hyphens
' and enclosed by curly braces. For example: {B09DE715-87C1-11D1-8BE3-0000F8754DA1}
' If pwszGuidText is omited, AfxGuid generates a new unique guid.
' Remarks: I have need to call the UuidCreate function dynamically because, at the time of
' writing, the library for the RPCRT4.DLL seems broken and the linker fails.
' ========================================================================================
PRIVATE FUNCTION Afx_OleConDispatch.GuidFromStr (BYVAL pwszGuidText AS WSTRING PTR = NULL) AS GUID
   DIM rguid AS GUID
   IF pwszGuidText = NULL THEN
      ' // Generate a new guid
      DIM AS ANY PTR pLib = DyLibLoad("RPCRT4.DLL")
      IF pLib  THEN
         DIM pProc AS FUNCTION (BYVAL Uuid AS UUID PTR) AS RPC_STATUS
         pProc = DyLibSymbol(pLib, "UuidCreate")
         IF pProc THEN pProc(@rguid)
         DyLibFree(pLib)
      END IF
   ELSE
      CLSIDFromString(pwszGuidText, @rGuid)
   END IF
   RETURN rguid
END FUNCTION
' ========================================================================================
