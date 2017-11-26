' ########################################################################################
' Microsoft Windows
' File: CAxHost_IDispatch.inc
' Contents: OLE Container class
' Implementation of the IDIspatch interface.
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2017 José Roca. Freeware. Use at your own risk.
' this CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once

' ########################################################################################
'                              *** CAxHost_IDispatch class ***
' ########################################################################################

' ========================================================================================
' CAxHost_OleConDispatch class
' ========================================================================================
TYPE CAxHost_IDispatch EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULONG
   DECLARE VIRTUAL FUNCTION GetTypeInfoCount (BYVAL pctinfo AS UINT PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION GetTypeInfo (BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION GetIDsOfNames (BYVAL riid AS CONST IID CONST PTR, BYVAL rgszNames AS LPOLESTR PTR, BYVAL cNames AS UINT, BYVAL lcid AS LCID, BYVAL rgDispId AS DISPID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION Invoke (BYVAL dispIdMember AS DISPID, BYVAL riid AS CONST IID CONST PTR, BYVAL lcid AS LCID, BYVAL wFlags AS WORD, BYVAL pDispParams AS DISPPARAMS PTR, BYVAL pVarResult AS VARIANT PTR, BYVAL pExcepInfo AS EXCEPINFO PTR, BYVAL puArgErr AS UINT PTR) AS HRESULT
Public:
   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL pAxHost AS CAxHost PTR)
   DECLARE DESTRUCTOR

Private:
   cRef AS ULONG = 0
   m_pAxHost AS CAxHost PTR

END TYPE
' ========================================================================================

' ========================================================================================
' Constructors
' ========================================================================================
CONSTRUCTOR CAxHost_IDispatch
   CAXH_DP("CAxHost_IDispatch Default constructor")
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
CONSTRUCTOR CAxHost_IDispatch (BYVAL pAxHost AS CAxHost PTR)
   CAXH_DP("CAxHost_IDispatch constructor pAxHost = " & WSTR(pAxHost))
   m_pAxHost = pAxHost
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Destructor
' ========================================================================================
DESTRUCTOR CAxHost_IDispatch
   CAXH_DP("### CAxHost_IDispatch destructor")
   m_pAxHost->m_pIDispatchImpl = NULL
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
FUNCTION CAxHost_IDispatch.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   IF m_pAxHost = NULL THEN RETURN E_POINTER
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   RETURN m_pAxHost->QueryInterface(riid, ppvObj)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Increments the reference count for an interface on an object. This method should be called
' for every new copy of a pointer to an interface on an object.
' ========================================================================================
FUNCTION CAxHost_IDispatch.AddRef () AS ULONG
   this.cRef += 1
   CAXH_DP("CAxHost_IDispatch AddRef = " & WSTR(this.cRef))
   RETURN this.cRef
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
FUNCTION CAxHost_IDispatch.Release () AS DWORD
   this.cRef -= 1
   CAXH_DP("CAxHost_IDispatch Release = " & WSTR(this.cRef))
   FUNCTION = this.cRef
   IF this.cRef = 0 THEN DELETE @this
END FUNCTION
' ========================================================================================

' ========================================================================================
' Not required. This method is necessary only for containers that support nonstandard
' ambient properties.
' ========================================================================================
FUNCTION CAxHost_IDispatch.GetTypeInfoCount (BYVAL pctInfo AS UINT PTR) AS HRESULT
   CAXH_DP("CAxHost_IDispatch GetTypeInfoCount")
   *pctInfo = 0
   RETURN E_NOTIMPL
END FUNCTION
' ========================================================================================

' ========================================================================================
' Not required. This method is necessary only for containers that support nonstandard
' ambient properties.
' ========================================================================================
FUNCTION CAxHost_IDispatch.GetTypeInfo (BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   CAXH_DP("CAxHost_IDispatch GetTypeInfo")
   RETURN E_NOTIMPL
END FUNCTION
' ========================================================================================

' ========================================================================================
' Not required. This method is necessary only for containers that support nonstandard
' ambient properties.
' ========================================================================================
FUNCTION CAxHost_IDispatch.GetIDsOfNames (BYVAL riid AS CONST IID CONST PTR, BYVAL rgszNames AS LPOLESTR PTR, BYVAL cNames AS UINT, BYVAL lcid AS LCID, BYVAL rgDispId AS DISPID PTR) AS HRESULT
   CAXH_DP("CAxHost_IDispatch GetIDsOfNames")
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
FUNCTION CAxHost_IDispatch.Invoke (BYVAL dispIdMember AS DISPID, BYVAL riid AS CONST IID CONST PTR, BYVAL lcid AS LCID, BYVAL wFlags AS WORD, BYVAL pDispParams AS DISPPARAMS PTR, BYVAL pVarResult AS VARIANT PTR, BYVAL pExcepInfo AS EXCEPINFO PTR, BYVAL puArgErr AS UINT PTR) AS HRESULT
   CAXH_DP("CAxHost_IDispatch Invoke - dispid = " & WSTR(dispIdMember))

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
         pvarResult->lVal = m_pAxHost->m_AmbientDisp.BackColor
         RETURN NOERROR
      CASE -703   ' DISPID_AMBIENT_FONT
         IF m_pAxHost->m_AmbientDisp.Font THEN
            ' Incremenet the reference count
            AfxSafeAddRef(m_pAxHost->m_AmbientDisp.Font)
            ' Return a reference to the ambient font
            pvarResult->vt = VT_DISPATCH
            pvarResult->pdispVal = CAST(IDispatch PTR, CAST(ULONG_PTR, m_pAxHost->m_AmbientDisp.Font))
            RETURN NOERROR
         END IF
      CASE DISPID_AMBIENT_FORECOLOR ' -704
         pvarResult->vt = VT_UI4
         pvarResult->lVal = m_pAxHost->m_AmbientDisp.Forecolor
         RETURN NOERROR
      CASE DISPID_AMBIENT_LOCALEID ' -705
         pvarResult->vt = VT_I4
         pvarResult->lVal = m_pAxHost->m_AmbientDisp.LocaleID
         RETURN NOERROR
      CASE DISPID_AMBIENT_MESSAGEREFLECT ' -706
         pvarResult->vt = VT_BOOL
         pvarResult->boolVal = 0
         RETURN NOERROR
      CASE DISPID_AMBIENT_USERMODE   ' -709
         pvarResult->vt = VT_BOOL
         pvarResult->boolVal = -1
         RETURN NOERROR
      CASE DISPID_AMBIENT_UIDEAD   ' -710
         pvarResult->vt = VT_BOOL
         pvarResult->boolVal = m_pAxHost->m_AmbientDisp.UIDead
         RETURN NOERROR
      CASE DISPID_AMBIENT_SHOWGRABHANDLES ' -711
         pvarResult->vt = VT_BOOL
         pvarResult->boolVal = 0
         RETURN NOERROR
      CASE DISPID_AMBIENT_SHOWHATCHING ' -712
         pvarResult->vt = VT_BOOL
         pvarResult->boolVal = 0
         RETURN NOERROR
      CASE DISPID_AMBIENT_DISPLAYASDEFAULT  ' -713
         pvarResult->vt = VT_BOOL
         pvarResult->boolVal = m_pAxHost->m_AmbientDisp.DisplayAsDefault
         RETURN NOERROR
      CASE DISPID_AMBIENT_SUPPORTSMNEMONICS  ' -714
         pvarResult->vt = VT_BOOL
         pvarResult->boolVal = m_pAxHost->m_AmbientDisp.SupportMnemonics
         RETURN NOERROR
      CASE -5501 ' DISPID_AMBIENT_OFFLINEIFNOTCONNECTED
         pvarResult->vt = VT_BOOL
         pvarResult->boolVal = m_pAxHost->m_AmbientDisp.OffLineIfNoConnected
         RETURN NOERROR
      CASE -5502   '  DISPID_AMBIENT_SILENT
         pvarResult->vt = VT_BOOL
         pvarResult->boolVal = m_pAxHost->m_AmbientDisp.Silent
         RETURN NOERROR
      CASE -5512   ' DISPID_AMBIENT_DLCONTROL
         IF m_pAxHost->m_AmbientDisp.DlCtFlags THEN
            pvarResult->vt = VT_I4
            pvarResult->lVal = m_pAxHost->m_AmbientDisp.DlCtFlags
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

