' ########################################################################################
' Microsoft Windows
' File: CAxHost.inc
' Contents: OLE Container class
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2017 José Roca. Freeware. Use at your own risk.
' this CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "Afx/AfxCOM.inc"
#include once "Afx/CWindow.inc"
#include once "Afx/AfxExDisp.bi"
#include once "Afx/CAxHost/CAxHost.bi"
USING Afx

NAMESPACE Afx

' // Forward references
TYPE CAxHost_ AS CAxHost
TYPE CAxHost_IDispatch_ AS CAxHost_IDispatch
TYPE CAxHost_IOleClientSite_ AS CAxHost_IOleClientSite
TYPE CAxHost_IOleControlSite_ AS CAxHost_IOleControlSite
TYPE CAxHost_IOleInPlaceFrame_ AS CAxHost_IOleInPlaceFrame
TYPE CAxHost_IOleInPlaceSite_ AS CAxHost_IOleInPlaceSite
TYPE CAxHost_IServiceProvider_ AS CAxHost_IServiceProvider

' The class name for all Visual Basic intrinsic controls
'CONST IID_IVBControlClass = "{FCFB3D25-A0FA-1068-A738-08002B3371B5}"
' Visual Basic only guids
'CONST IID_IVBGetControl  = "{40A050A0-3C31-101B-A82E-08002B2B2337}"
'CONST CATID_VBGetControl = "{02496841-3AC4-11CF-87B9-00AA006C8166}"
'CONST IID_IGetOleObject = "{8A701DA0-4FEB-101B-A82E-08002B2B2337}"
'CONST IID_IVBFormat  = "{9849FD60-3768-101B-8D72-AE6164FFE3CF}"
'CONST CATID_VBFormat = "{02496840-3AC4-11CF-87B9-00AA006C8166}"
'CONST IID_IGetVBAObject = "{91733A60-3F4C-101B-A3F6-00AA0034E4E9}"
'CONST CLSID_DataObject = "{2334D2B2-713E-11CF-8AE5-00AA00C00905}"
'CONST CLSID_DataObjectFiles = "{2334D2B4-713E-11CF-8AE5-00AA00C00905}"
'CONST IID_IVBDataObjectFiles = "{2334D2B3-713E-11CF-8AE5-00AA00C00905}"
'CONST IID_IVBDataObject = "{2334D2B1-713E-11CF-8AE5-00AA00C00905}"

CONST DLCTL_DLIMAGES                  = &h00000010
CONST DLCTL_VIDEOS                    = &h00000020
CONST DLCTL_BGSOUNDS                  = &h00000040
CONST DLCTL_NO_SCRIPTS                = &h00000080
CONST DLCTL_NO_JAVA                   = &h00000100
CONST DLCTL_NO_RUNACTIVEXCTLS         = &h00000200
CONST DLCTL_NO_DLACTIVEXCTLS          = &h00000400
CONST DLCTL_DOWNLOADONLY              = &h00000800
CONST DLCTL_NO_FRAMEDOWNLOAD          = &h00001000
CONST DLCTL_RESYNCHRONIZE             = &h00002000
CONST DLCTL_PRAGMA_NO_CACHE           = &h00004000
CONST DLCTL_NO_BEHAVIORS              = &h00008000
CONST DLCTL_NO_METACHARSET            = &h00010000
CONST DLCTL_URL_ENCODING_DISABLE_UTF8 = &h00020000
CONST DLCTL_URL_ENCODING_ENABLE_UTF8  = &h00040000
CONST DLCTL_NOFRAMES                  = &h00080000
CONST DLCTL_FORCEOFFLINE              = &h10000000
CONST DLCTL_NO_CLIENTPULL             = &h20000000
CONST DLCTL_SILENT                    = &h40000000
CONST DLCTL_OFFLINEIFNOTCONNECTED     = &h80000000
CONST DLCTL_OFFLINE                   = DLCTL_OFFLINEIFNOTCONNECTED

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

' ########################################################################################
'                                *** CAXHost class ***
' ########################################################################################

' ========================================================================================
' CAxHost class
' Implements an OLE container AS a Windows custom control.
' ========================================================================================
TYPE CAxHost

Public:
   m_bUninitOLE AS BOOLEAN                                          ' // Ole initialization flag
   m_hwndContainer            AS HWND                               ' // Window handle
   m_pOcx                     AS IDispatch PTR                      ' // Address of a pointer to the ActiveX's Vtable
   m_hFont                    AS HFONT                              ' // Window font
   m_dwMiscStatus             AS DWORD                              ' // Status bitwise flags
   m_bLocked                  AS BOOLEAN                            ' // Container locked
   m_bInPlaceActive           AS BOOLEAN                            ' // The control is in-place active
   m_bUIActive                AS BOOLEAN                            ' // The control is UI Active
   m_bHaveFocus               AS BOOLEAN                            ' // The control has focus
   m_rxRatio                  AS SINGLE                             ' // DPI horizontal ratio
   m_ryRatio                  AS SINGLE                             ' // DPI vertical ratio
   m_wszProgID                AS WSTRING * 260                      ' // ActiveX ProgID
   m_wszLibName               AS WSTRING * 260                      ' // ActiveX path
   m_wszLicKey                AS WSTRING * 260                      ' // ActiveX license key
   m_rclsid                   AS CLSID                              ' // ActiveX ClsId
   m_riid                     AS IID                                ' // ActiveX IID
   m_bIsRegFree               AS BOOLEAN                            ' // Register free OCX
   m_bIsLicensed              AS BOOLEAN                            ' // Licensed OCX
   m_dwCookie                 AS DWORD                              ' // Cookie for Unadvise
   m_riidEvt                  AS IID PTR                            ' // IID of the events interface
   m_AmbientDisp              AS CAXHOST_AMBIENTDISP                ' // Ambient dispatch properties and flags

   ' *** Implemented interfaces ***
   m_pIDispatchImpl           AS CAxHost_IDispatch_ PTR             ' // IDispatch interface (implemented)
   m_pIOleClientSiteImpl      AS CAxHost_IOleClientSite_ PTR        ' // IOleClientSite interface (implemented)
   m_pIOleControlSiteImpl     AS CAxHost_IOleControlSite_ PTR       ' // IOleControlSite interface (implemented)
   m_pIOleInPlaceFrameImpl    AS CAxHost_IOleInPlaceFrame_ PTR      ' // IOleInPlaceFrame interface (implemented)
   m_pIOleInPlaceSiteImpl     AS CAxHost_IOleInPlaceSite_ PTR       ' // IOleInPlaceSite interface (implemented)
   m_pIServiceProviderImpl    AS CAxHost_IServiceProvider_ PTR      ' // IServiceProvider interface (implemented)

   ' *** Control interfaces ***
   m_pIOleObject              AS Afx_IOleObject PTR                 ' // IOleObject interface
   m_pIOleInPlaceObject       AS Afx_IOleInPlaceObject PTR          ' // IOleInPlaceObject interface
   m_pIOleInPlaceActiveObject AS Afx_IOleInPlaceActiveObject PTR    ' // IOleInPlaceActiveObject interface

Protected:
   DECLARE FUNCTION RegisterClass () AS ATOM
   DECLARE STATIC FUNCTION CAxHostProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT
   DECLARE SUB CreateContainer (BYVAL pWindow AS CWindow PTR, BYVAL cID AS LONG_PTR, BYREF wszProgID AS WSTRING, _
      BYREF wszLibName AS WSTRING, BYREF rclsid AS const CLSID, BYREF riid AS const IID, BYREF wszLicKey AS WSTRING, _
      BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
      BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, BYVAL pAmbientDisp AS CAXHOST_AMBIENTDISP PTR = NULL)

Public:
   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL pWindow AS CWindow PTR, BYVAL cID AS LONG_PTR, BYREF wszLibName AS WSTRING, _
      BYREF rclsid AS const CLSID, BYREF riid AS const IID, BYREF wszLicKey AS WSTRING, _
      BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
      BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, BYVAL pAmbientDisp AS CAXHOST_AMBIENTDISP PTR = NULL)
   DECLARE CONSTRUCTOR (BYVAL pWindow AS CWindow PTR, BYVAL cID AS LONG_PTR, BYREF wszLibName AS WSTRING, _
      BYREF rclsid AS const CLSID, BYREF riid AS const IID, BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, _
      BYVAL nHeight AS LONG = 0, BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, BYVAL pAmbientDisp AS CAXHOST_AMBIENTDISP PTR = NULL)
   DECLARE CONSTRUCTOR (BYVAL pWindow AS CWindow PTR, BYVAL cID AS LONG_PTR, BYREF wszProgID AS WSTRING, _
      BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
      BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, BYVAL pAmbientDisp AS CAXHOST_AMBIENTDISP PTR = NULL)
   DECLARE CONSTRUCTOR (BYVAL pWindow AS CWindow PTR, BYVAL cID AS LONG_PTR, BYREF wszProgID AS WSTRING, _
      BYREF wszLicKey AS WSTRING, BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
      BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, BYVAL pAmbientDisp AS CAXHOST_AMBIENTDISP PTR = NULL)
   DECLARE CONSTRUCTOR (BYVAL pWindow AS CWindow PTR, BYVAL cID AS LONG_PTR, BYREF classID AS CONST CLSID, _
      BYREF riid AS CONST IID, BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, _
      BYVAL nHeight AS LONG = 0, BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, BYVAL pAmbientDisp AS CAXHOST_AMBIENTDISP PTR = NULL)
   DECLARE CONSTRUCTOR (BYVAL pWindow AS CWindow PTR, BYVAL cID AS LONG_PTR, BYREF classID AS CONST CLSID, _
      BYREF riid AS CONST IID, BYREF wszLicKey AS WSTRING, BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, _
      BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, BYVAL dwStyle AS DWORD = 0, _
      BYVAL dwExStyle AS DWORD = 0, BYVAL pAmbientDisp AS CAXHOST_AMBIENTDISP PTR = NULL)
   DECLARE FUNCTION CreateObject (BYREF wszProgID AS WSTRING) AS IDispatch PTR
   DECLARE FUNCTION CreateObject (BYREF wszProgID AS WSTRING, BYREF wszLicKey AS WSTRING) AS IDispatch PTR
   DECLARE FUNCTION CreateObject (BYREF classID AS CONST CLSID, BYREF riid AS CONST IID) AS IDispatch PTR
   DECLARE FUNCTION CreateObject (BYREF classID AS CONST CLSID, BYREF riid AS CONST IID, BYREF wszLicKey AS WSTRING) AS IDispatch PTR
   DECLARE FUNCTION CreateObject (BYREF wszLibName AS CONST WSTRING, BYREF rclsid AS CONST CLSID, _
      BYREF riid AS CONST IID, BYREF wszLicKey AS WSTRING = "") AS IDispatch PTR
   DECLARE FUNCTION hWindow () AS HANDLE
   DECLARE FUNCTION CtrlID () AS LONG
   DECLARE SUB PixelsToHiMetric (BYVAL SizeInPix AS SIZEL PTR, BYVAL SizeInHiMetric AS SIZEL PTR)
   DECLARE FUNCTION OcxDispPtr () AS IDispatch PTR
   DECLARE FUNCTION OcxDispObj () AS IDispatch PTR
   DECLARE FUNCTION Advise (BYVAL pEvtObj AS IDispatch PTR, BYVAL riid AS IID PTR) AS HRESULT
   DECLARE FUNCTION Advise (BYVAL pEvtObj AS IDispatch PTR, BYREF riid AS CONST IID) AS HRESULT
   DECLARE FUNCTION Advise (BYVAL pEvtObj AS IDispatch PTR, BYREF riid AS IID) AS HRESULT
   DECLARE FUNCTION Advise (BYVAL pEvtObj AS Afx_IDispatch PTR, BYVAL riid AS IID PTR) AS HRESULT
   DECLARE FUNCTION Advise (BYVAL pEvtObj AS Afx_IDispatch PTR, BYREF riid AS CONST IID) AS HRESULT
   DECLARE FUNCTION Advise (BYVAL pEvtObj AS Afx_IDispatch PTR, BYREF riid AS IID) AS HRESULT
   DECLARE FUNCTION Unadvise () AS HRESULT
   DECLARE FUNCTION OleCreateFontDisp (BYREF wszFontName AS WSTRING, BYVAL cySize AS LONGLONG, _
           BYVAL fWeight AS SHORT = FW_NORMAL, BYVAL fItalic AS WINBOOL = FALSE, BYVAL fUnderline AS WINBOOL = FALSE, _
           BYVAL fStrikethrough AS WINBOOL = FALSE, BYVAL fCharset AS SHORT = DEFAULT_CHARSET) AS IFontDisp PTR
   DECLARE FUNCTION OleCreateFont (BYREF wszFontName AS WSTRING, BYVAL cySize AS LONGLONG, _
           BYVAL fWeight AS SHORT = FW_NORMAL, BYVAL fItalic AS WINBOOL = FALSE, BYVAL fUnderline AS WINBOOL = FALSE, _
           BYVAL fStrikethrough AS WINBOOL = FALSE, BYVAL fCharset AS SHORT = DEFAULT_CHARSET) AS IFont PTR
   DECLARE FUNCTION GuidFromStr (BYVAL pwszGuidText AS WSTRING PTR = NULL) AS GUID
   DECLARE FUNCTION GuidText (BYVAL classID AS CLSID PTR) AS STRING
   DECLARE FUNCTION GuidText (BYVAL riid AS REFIID) AS STRING
   DECLARE FUNCTION GuidText (BYVAL classID AS CLSID) AS STRING
   DECLARE FUNCTION GetInterfaceName (BYVAL riid AS REFIID) AS STRING
   DECLARE FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT

END TYPE
' ========================================================================================

' ========================================================================================
' Include the implemented interfaces
' We need to include them here because they reference CAxHost.
' ========================================================================================
#include once "CAxHost_IDispatch.inc"
#include once "CAxHost_IOleClientSite.inc"
#include once "CAxHost_IOleControlSite.inc"
#include once "CAxHost_IOleInPlaceFrame.inc"
#include once "CAxHost_IOleInPlaceSite.inc"
#include once "CAxHost_IServiceProvider.inc"
' ========================================================================================

' ========================================================================================
' Constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CAxHost
   CAXH_DP("CaxHost CONSTRUCTOR")
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Destructor
' ========================================================================================
PRIVATE DESTRUCTOR CAxHost
   CAXH_DP("--- CAXHOST DESTRUCTOR - BEGIN ---")

   ' // Release de IFontDisp object
   DIM cRef AS ULONG = AfxSafeRelease(m_AmbientDisp.Font)
   CAXH_DP("Release IFontDisp - cRef = " & WSTR(cRef))

   ' // Close the control
   IF m_pIOleObject THEN
      CAXH_DP("--- CAXHOST DESTRUCTOR - CLOSE - BEGIN ---")
      m_pIOleObject->Close(OLECLOSE_NOSAVE)
      CAXH_DP("--- CAXHOST DESTRUCTOR - CLOSE - END ---")
      m_pIOleObject->SetClientSite(NULL)
      CAXH_DP("--- CAXHOST DESTRUCTOR - SetClientSite NULL ---")
   END IF

   ' // Release IOleInPlaceActiveObject
   cRef = AfxSafeRelease(m_pIOleInPlaceActiveObject)
   CAXH_DP("Release IOleInPlaceActiveObject - cRef (must be 0) = " & WSTR(cRef))

   ' // Release the IOleObject interface
   ' // The returned reference count must be at least 1 because the
   ' // ActiveX control holds a reference to it.
   cRef = AfxSafeRelease(m_pIOleObject)
   CAXH_DP("Release IOleObject - cRef = " & WSTR(cRef))

   ' // WORKAROUND: The DHTML Edit Control doesn't release this interface
   IF this.m_pIOleControlSiteImpl THEN AfxSafeRelease(this.m_pIOleControlSiteImpl)

   ' // Release the IOleClientSite interface
   IF this.m_pIOleClientSiteImpl THEN AfxSafeRelease(this.m_pIOleClientSiteImpl)

   ' // Delete the IServiceProvider class
   IF this.m_pIServiceProviderImpl THEN DELETE this.m_pIServiceProviderImpl

   ' // Release the ActiveX control
   cRef = AfxSafeRelease(m_pOcx)
   CAXH_DP("Release OCX - cRef = " & WSTR(cRef))

   ' // Checks for memory leaks (if a pointer <> 0 then memory leak)
   CAXH_DP("*** IDispatch = " & WSTR(this.m_pIDispatchImpl))
   CAXH_DP("*** IOleClientSite = " & WSTR(this.m_pIOleClientSiteImpl))
   CAXH_DP("*** IOleControlSite = " & WSTR(this.m_pIOleControlSiteImpl))
   CAXH_DP("*** IOleInPlaceFrame = " & WSTR(this.m_pIOleInPlaceFrameImpl))
   CAXH_DP("*** IOleInPlaceSite = " & WSTR(this.m_pIOleInPlaceSiteImpl))
   CAXH_DP("*** IServiceProvider = " & WSTR(this.m_pIServiceProviderImpl))
   CAXH_DP("*** IOleInPlaceActiveObject = " & WSTR(this.m_pIOleInPlaceActiveObject))
   CAXH_DP("*** IOleInPlaceObject = " & WSTR(this.m_pIOleInPlaceObject))
   CAXH_DP("*** IOleObject = " & WSTR(this.m_pIOleObject))
   CAXH_DP("*** bInPlaceActive = " & WSTR(this.m_bInPlaceActive))

   ' // Free the COM library
   CAXH_DP("DESTRUCTOR CAxHost - m_bUninitOLE = " & WSTR(m_bUninitOLE))
   IF m_bUninitOLE THEN OleUninitialize

   CAXH_DP("--- CAXHOST DESTRUCTOR - END ---")

END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Registers the class
' ========================================================================================
PRIVATE FUNCTION CAxHost.RegisterClass () AS ATOM
   DIM wAtom AS ATOM
   DIM wcexw AS WNDCLASSEXW
   DIM wszClassName AS WSTRING * 260 = CAXHOST_CLASSNAME
   IF .GetClassInfoExW(.GetModuleHandleW(NULL), @wszClassName, @wcexw) = 0 THEN
      ' // Fill the WNDCLASSEXW structure
      WITH wcexw
         .cbSize        = SIZEOF(wcexw)
         .style         = CS_DBLCLKS OR CS_HREDRAW OR CS_VREDRAW
         .lpfnWndProc   = @CAxHostProc
         .cbClsExtra    = 0
         .cbWndExtra    = SIZEOF(HANDLE)
         .hInstance     = ..GetModuleHandleW(NULL)
         .hCursor       = ..LoadCursorW(NULL, CAST(LPCWSTR, IDC_ARROW))
         .hbrBackground = NULL
         .lpszMenuName  = NULL
         .lpszClassName = @wszClassName
         .hIcon         = NULL
         .hIconSm       = NULL
      END WITH
      wAtom = .RegisterClassExW(@wcexw)
   END IF
   RETURN wAtom
END FUNCTION
' ========================================================================================

' ========================================================================================
' Create container window
' ========================================================================================
PRIVATE SUB CAxHost.CreateContainer (BYVAL pWindow AS CWindow PTR, BYVAL cID AS LONG_PTR, BYREF wszProgID AS WSTRING, _
   BYREF wszLibName AS WSTRING, BYREF rclsid AS const CLSID, BYREF riid AS const IID, BYREF wszLicKey AS WSTRING, _
   BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
   BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, BYVAL pAmbientDisp AS CAXHOST_AMBIENTDISP PTR = NULL)

   CAXH_DP("CONSTRUCTOR CAxHost")
   ' // Initialize the COM library
   DIM hr AS HRESULT = OleInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitOLE = TRUE
   CAXH_DP("CONSTRUCTOR CAxHost - OleInitialize hr = " & WSTR(hr))
   ' // Register the class
   DIM wszClassName AS WSTRING * 260 = CAXHOST_CLASSNAME
   this.RegisterClass
   ' // Store values
   IF pAmbientDisp THEN
      WITH m_AmbientDisp
         .Font = pAmbientDisp->Font
         .BackColor = pAmbientDisp->BackColor
         .ForeColor = pAmbientDisp->ForeColor
         .LocaleID = pAmbientDisp->LocaleID
         .Silent = pAmbientDisp->Silent
         .UIDead = pAmbientDisp->UIDead
         .DisplayAsDefault = pAmbientDisp->DisplayAsDefault
         .SupportMnemonics = pAmbientDisp->SupportMnemonics
         .DlCtFlags = pAmbientDisp->DlCtFlags
      END WITH
   END IF
   m_rxRatio = pWindow->rxRatio
   m_ryRatio = pWindow->ryRatio
   m_wszProgID = wszProgID     ' // ActiveX ProgID
   m_wszLibName = wszLibName   ' // ActiveX path
   m_wszLicKey = wszLicKey     ' // Active license key
   m_rclsid = rclsid           ' // ActiveX CLSID
   m_riid = riid               ' // ActiveX IID
   IF LEN(wszLibName) THEN m_bIsRegFree = TRUE
   IF LEN(wszLicKey) THEN m_bIsLicensed = TRUE

   ' // Create the control
   IF dwStyle = 0 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP
   IF pWindow THEN m_hwndContainer = pWindow->AddControl(wszClassName, pWindow->hWindow, cID, "", _
      x, y, nWidth, nHeight, dwStyle, dwExStyle, CAST(LONG_PTR, @this))
   IF m_hwndContainer THEN
      .SetWindowLongPtrW m_hwndContainer, 0, CAST(LONG_PTR, @this)
      ' // Set the same font used by the parent
      DIM lfw AS LOGFONTW
      IF pWindow->Font THEN
         IF .GetObjectW(pWindow->Font, SIZEOF(lfw), @lfw) THEN m_hFont = CreateFontIndirectW(@lfw)
      END IF
   END IF

END SUB
' ========================================================================================

' ========================================================================================
' COleCon class constructors
' ========================================================================================
PRIVATE CONSTRUCTOR CAxHost (BYVAL pWindow AS CWindow PTR, BYVAL cID AS LONG_PTR, BYREF wszLibName AS WSTRING, _
BYREF rclsid AS const CLSID, BYREF riid AS const IID, BYREF wszLicKey AS WSTRING, _
BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, BYVAL pAmbientDisp AS CAXHOST_AMBIENTDISP PTR = NULL)
   this.CreateContainer(pWindow, cID, "", wszLibName, rclsid, riid, wszLicKey, x, y, nWidth, nHeight, dwStyle, dwExStyle, pAmbientDisp)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CAxHost (BYVAL pWindow AS CWindow PTR, BYVAL cID AS LONG_PTR, BYREF wszLibName AS WSTRING, _
BYREF rclsid AS const CLSID, BYREF riid AS const IID, BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, _
BYVAL nHeight AS LONG = 0, BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, BYVAL pAmbientDisp AS CAXHOST_AMBIENTDISP PTR = NULL)
   this.CreateContainer(pWindow, cID, "", wszLibName, rclsid, riid, "", x, y, nWidth, nHeight, dwStyle, dwExStyle, pAmbientDisp)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CAxHost (BYVAL pWindow AS CWindow PTR, BYVAL cID AS LONG_PTR, BYREF wszProgID AS WSTRING, _
BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, BYVAL pAmbientDisp AS CAXHOST_AMBIENTDISP PTR = NULL)
   this.CreateContainer(pWindow, cID, wszProgID, "", GUID_NULL, GUID_NULL, "", x, y, nWidth, nHeight, dwStyle, dwExStyle, pAmbientDisp)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CAxHost (BYVAL pWindow AS CWindow PTR, BYVAL cID AS LONG_PTR, BYREF wszProgID AS WSTRING, _
BYREF wszLicKey AS WSTRING, BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, BYVAL pAmbientDisp AS CAXHOST_AMBIENTDISP PTR = NULL)
   this.CreateContainer(pWindow, cID, wszProgID, "", GUID_NULL, GUID_NULL, wszLicKey, x, y, nWidth, nHeight, dwStyle, dwExStyle, pAmbientDisp)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CAxHost (BYVAL pWindow AS CWindow PTR, BYVAL cID AS LONG_PTR, BYREF classID AS CONST CLSID, _
BYREF riid AS CONST IID, BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, _
BYVAL nHeight AS LONG = 0, BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, BYVAL pAmbientDisp AS CAXHOST_AMBIENTDISP PTR = NULL)
   this.CreateContainer(pWindow, cID, "", "", classID, riid, "", x, y, nWidth, nHeight, dwStyle, dwExStyle, pAmbientDisp)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CAxHost (BYVAL pWindow AS CWindow PTR, BYVAL cID AS LONG_PTR, BYREF classID AS CONST CLSID, _
BYREF riid AS CONST IID, BYREF wszLicKey AS WSTRING, BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, _
BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, BYVAL dwStyle AS DWORD = 0, _
BYVAL dwExStyle AS DWORD = 0, BYVAL pAmbientDisp AS CAXHOST_AMBIENTDISP PTR = NULL)
   this.CreateContainer(pWindow, cID, "", "", classID, riid, wszLicKey, x, y, nWidth, nHeight, dwStyle, dwExStyle, pAmbientDisp)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the handle of the container's window
' ========================================================================================
PRIVATE FUNCTION CAxHost.hWindow () AS HANDLE
   FUNCTION = m_hwndContainer
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the identifier of the container's window
' ========================================================================================
PRIVATE FUNCTION CAxHost.CtrlID () AS LONG
   FUNCTION = GetDLgCtrlID(m_hwndContainer)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a reference to the hosted ActiveX control dispatch interface.
' Don't call IUnknown_Release.
' ========================================================================================
PRIVATE FUNCTION CAxHost.OcxDispPtr () AS IDispatch PTR
   RETURN m_pOcx
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a counted reference to the hosted ActiveX control dispatch interface.
' You must call IUnknown_Release when no longer need it.
' ========================================================================================
PRIVATE FUNCTION CAxHost.OcxDispObj () AS IDispatch PTR
   IF m_pOcx <> NULL THEN
      m_pOcx->lpvtbl->AddRef(m_pOcx)
      RETURN m_pOcx
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Establishes a connection between the connection point object and the client's sink.
' ========================================================================================
PRIVATE FUNCTION CAxHost.Advise (BYVAL pEvtObj AS IDispatch PTR, BYVAL riid AS IID PTR) AS HRESULT
   CAXH_DP("BEGIN CAxHost.Advise")
   IF pEvtObj = NULL THEN RETURN E_POINTER
   m_riidEvt = riid
   ' // Query for the IConnectionPointContainer interface
   DIM pCPC AS IConnectionPointContainer PTR
   DIM hr AS HRESULT = IUnknown_QueryInterface(m_pOcx, @IID_IConnectionPointContainer, @pCPC)
   IF hr <> S_OK OR pCPC = NULL THEN RETURN hr
   ' // Query for the IConnectionPoint interface
   DIM pCP AS IConnectionPoint PTR
   hr = pCPC->lpvtbl->FindConnectionPoint(pCPC, riid, @pCP)
   IF hr <> S_OK OR pCP = NULL THEN
      AfxSafeRelease(pCPC)
      RETURN hr
   END IF
   ' // Terminates the advisory connection previously established between a connection point object and a client's sink.
   IF m_dwCookie THEN hr = pCP->lpvtbl->Unadvise(pCP, m_dwCookie)
   m_dwCookie = 0
   ' // Establishes a connection between a connection point object and the client's sink.
   hr = pCP->lpvtbl->Advise(pCP, cast(IUnknown PTR, pEvtObj), @m_dwCookie)
   AfxSafeRelease(pCPC)
   AfxSafeRelease(pCP)
   CAXH_DP("END CAxHost.Advise")
   RETURN hr
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CAxHost.Advise (BYVAL pEvtObj AS IDispatch PTR, BYREF riid AS CONST IID) AS HRESULT
   RETURN this.Advise(pEvtObj, cast(IID PTR, @riid))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CAxHost.Advise (BYVAL pEvtObj AS IDispatch PTR, BYREF riid AS IID) AS HRESULT
   RETURN this.Advise(pEvtObj, cast(IID PTR, @riid))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CAxHost.Advise (BYVAL pEvtObj AS Afx_IDispatch PTR, BYVAL riid AS IID PTR) AS HRESULT
   RETURN this.Advise(cast(IDispatch PTR, cast(ULONG_PTR, pEvtObj)), riid)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CAxHost.Advise (BYVAL pEvtObj AS Afx_IDispatch PTR, BYREF riid AS CONST IID) AS HRESULT
   RETURN this.Advise(cast(IDispatch PTR, cast(ULONG_PTR, pEvtObj)), cast(IID PTR, @riid))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CAxHost.Advise (BYVAL pEvtObj AS Afx_IDispatch PTR, BYREF riid AS IID) AS HRESULT
   RETURN this.Advise(cast(IDispatch PTR, cast(ULONG_PTR, pEvtObj)), cast(IID PTR, @riid))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Releases the events connection identified with the cookie returned by the Advise method.
' ========================================================================================
PRIVATE FUNCTION CAxHost.Unadvise () AS HRESULT
   CAXH_DP("BEGIN CAxHost.Unadvise")
   ' // Not a valid connection
   IF m_dwCookie = 0 THEN RETURN E_POINTER
   ' // Query for the IConnectionPointContainer interface
   DIM pCPC AS IConnectionPointContainer PTR
   DIM hr AS HRESULT = m_pOcx->lpvtbl->QueryInterface(m_pOcx, @IID_IConnectionPointContainer, @pCPC)
   IF hr <> S_OK OR pCPC = NULL THEN RETURN hr
   ' // Query for the IConnectionPoint interface
   DIM pCP AS IConnectionPoint PTR
   hr = pCPC->lpvtbl->FindConnectionPoint(pCPC, m_riidEvt, @pCP)
   IF hr <> S_OK OR pCP = NULL THEN
      AfxSafeRelease(pCPC)
      RETURN hr
   END IF
   ' // Terminates the advisory connection previously established between a connection point object and a client's sink.
   hr = pCP->lpvtbl->Unadvise(pCP, m_dwCookie)
   m_dwCookie = 0
   AfxSafeRelease(pCPC)
   AfxSafeRelease(pCP)
   CAXH_DP("END CAxHost.Unadvise")
   RETURN hr
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts from Pixel to HiMetric
' Note: Himetric is a scaling unit similar to twips used in computing. It is one
' thousandth of a centimeter and is independent of the screen resolution.
' HIMETRIC_PER_INCH = 2540   ' 1 inch = 2.54 mm
' ========================================================================================
PRIVATE SUB CAxHost.PixelsToHiMetric (BYVAL SizeInPix AS SIZEL PTR, BYVAL SizeInHiMetric AS SIZEL PTR)

   DIM nPixelsPerInchX AS LONG   ' // Pixels per logical inch along width
   DIM nPixelsPerInchY AS LONG   ' // Pixels per logical inch along height

   DIM hDCSCreen AS HDC = GetDC(NULL)
   IF hDCScreen = NULL THEN EXIT SUB
   nPixelsPerInchX = GetDeviceCaps(hDCScreen, LOGPIXELSX)
   nPixelsPerInchY = GetDeviceCaps(hDCScreen, LOGPIXELSY)
   ReleaseDC(NULL, hDCScreen)

   SizeInHiMetric->cx = MulDiv(SizeInPix->cx, 2540, nPixelsPerInchX)
   SizeInHiMetric->cy = MulDiv(SizeInPix->cy, 2540, nPixelsPerInchY)

END SUB
' ========================================================================================

' ========================================================================================
' Creates a standard IFontDisp object
' Parameters:
' - wszFontName: The typeface name, e.g. "Segoe UI"
' - cySize: The point size, e.g. 9.
' - fWeight: Initial weight of the font. If the weight is below 550 (the average of
'      FW_NORMAL, 400, and FW_BOLD, 700), then the Bold property is also initialized to
'      FALSE. If the weight is above 550, the Bold property is set to TRUE.
'      The following values are defined for convenience.
'      FW_DONTCARE (0), FW_THIN (100), FW_EXTRALIGHT (200), FW_ULTRALIGHT (200), FW_LIGHT (300),
'      FW_NORMAL (400), FW_REGULAR (400), FW_MEDIUM (500), FW_SEMIBOLD (600), FW_DEMIBOLD (600),
'      FW_BOLD (700), FW_EXTRABOLD (800), FW_ULTRABOLD (800), FW_HEAVY (900), FW_BLACK (900)
' - bItalic = Italic flag. CTRUE or FALSE
' - bUnderline = Underline flag. CTRUE or FALSE
' - bStrikeOut = StrikeOut flag. CTRUE or FALSE
' - bCharset = Charset.
'      The following values are predefined: ANSI_CHARSET, BALTIC_CHARSET, CHINESEBIG5_CHARSET,
'      DEFAULT_CHARSET, EASTEUROPE_CHARSET, GB2312_CHARSET, GREEK_CHARSET, HANGUL_CHARSET,
'      MAC_CHARSET, OEM_CHARSET, RUSSIAN_CHARSET, SHIFTJIS_CHARSET, SYMBOL_CHARSET, TURKISH_CHARSET,
'      VIETNAMESE_CHARSET, JOHAB_CHARSET (Korean language edition of Windows), ARABIC_CHARSET and
'      HEBREW_CHARSET (Middle East language edition of Windows), THAI_CHARSET (Thai language
'      edition of Windows).
'      The OEM_CHARSET value specifies a character set that is operating-system dependent.
'      DEFAULT_CHARSET is set to a value based on the current system locale. For example, when
'      the system locale is English (United States), it is set as ANSI_CHARSET.
'      Fonts with other character sets may exist in the operating system. If an application uses
'      a font with an unknown character set, it should not attempt to translate or interpret
'      strings that are rendered with that font.
'      This parameter is important in the font mapping process. To ensure consistent results,
'      specify a specific character set. If you specify a typeface name in the lfFaceName member,
'      make sure that the lfCharSet value matches the character set of the typeface specified in lfFaceName.
' Return value: A pointer to the object or NULL on failure.
' Remarks: The returned font must be destroyed calling the release method of the IFontDisp
' interface when no longer needed to prevent memory leaks.
' Usage examples:
' pFont = pAxHost.OleCreateFontDisp("MS Sans Serif", 8, FW_NORMAL, , , , DEFAULT_CHARSET)
' pFont = pAxHost.OleCreateFontDisp("Courier New", 10, FW_BOLD, , , , DEFAULT_CHARSET)
' pFont = pAxHost.OleCreateFontDisp("Marlett", 8, FW_NORMAL, , , , SYMBOL_CHARSET)
' ========================================================================================
PRIVATE FUNCTION CAxHost.OleCreateFontDisp ( _
   BYREF wszFontName AS WSTRING, _                  ' __in  Font name
   BYVAL cySize AS LONGLONG, _                      ' __in  Point size
   BYVAL fWeight AS SHORT = FW_NORMAL, _            ' __in  Weight (FW_NORMAL, etc.)
   BYVAL fItalic AS WINBOOL = FALSE, _              ' __in  Italic state
   BYVAL fUnderline AS WINBOOL = FALSE, _           ' __in  Underline state
   BYVAL fStrikethrough AS WINBOOL = FALSE, _       ' __in  Strikethrough state
   BYVAL fCharset AS SHORT = DEFAULT_CHARSET _      ' __in  Character set
   ) AS IFontDisp PTR

   CAXH_DP("OleCreateFontDisp " & wszFontName)
   DIM tf AS FONTDESC
   tf.cbSizeOfStruct = SIZEOF(FONTDESC)
   tf.lpstrName =  @wszFontName
   tf.cySize.int64 = cySize * 10000
   tf.sWeight = fWeight
   tf.sCharset = fCharset
   tf.fItalic = fItalic
   tf.fUnderline = fUnderline
   tf.fStrikethrough = fStrikethrough
   DIM pFont AS IFontDisp PTR
   OleCreateFontIndirect(@tf, @IID_IDispatch, @pFont)
   RETURN pFont

END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CAxHost.OleCreateFont ( _
   BYREF wszFontName AS WSTRING, _                  ' __in  Font name
   BYVAL cySize AS LONGLONG, _                      ' __in  Point size
   BYVAL fWeight AS SHORT = FW_NORMAL, _            ' __in  Weight (FW_NORMAL, etc.)
   BYVAL fItalic AS WINBOOL = FALSE, _              ' __in  Italic state
   BYVAL fUnderline AS WINBOOL = FALSE, _           ' __in  Underline state
   BYVAL fStrikethrough AS WINBOOL = FALSE, _       ' __in  Strikethrough state
   BYVAL fCharset AS SHORT = DEFAULT_CHARSET _      ' __in  Character set
   ) AS IFont PTR

   CAXH_DP("OleCreateFont " & wszFontName)
   DIM tf AS FONTDESC
   tf.cbSizeOfStruct = SIZEOF(FONTDESC)
   tf.lpstrName =  @wszFontName
   tf.cySize.int64 = cySize * 10000
   tf.sWeight = fWeight
   tf.sCharset = fCharset
   tf.fItalic = fItalic
   tf.fUnderline = fUnderline
   tf.fStrikethrough = fStrikethrough
   DIM pFont AS IFont PTR
   OleCreateFontIndirect(@tf, @IID_IUnknown, @pFont)
   RETURN pFont

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
PRIVATE FUNCTION CAxHost.GuidFromStr (BYVAL pwszGuidText AS WSTRING PTR = NULL) AS GUID
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

' ========================================================================================
' Returns a 38-byte human-readable guid string from a 16-byte GUID.
' ========================================================================================
PRIVATE FUNCTION CAxHost.GuidText (BYVAL classID AS CLSID PTR) AS STRING
   DIM pwsz AS WSTRING PTR
   StringFromCLSID(classID, CAST(LPOLESTR PTR, @pwsz))
   FUNCTION = *pwsz
   CoTaskMemFree(pwsz)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CAxHost.GuidText (BYVAL riid AS REFIID) AS STRING
   DIM pwsz AS WSTRING PTR
   StringFromIID(riid, CAST(LPOLESTR PTR, @pwsz))
   FUNCTION = *pwsz
   CoTaskMemFree(pwsz)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CAxHost.GuidText (BYVAL classID AS CLSID) AS STRING
   DIM pwsz AS WSTRING PTR
   StringFromCLSID(@classID, CAST(LPOLESTR PTR, @pwsz))
   FUNCTION = *pwsz
   CoTaskMemFree(pwsz)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates an instance of an object.
' Parameters:
' - wszProgID: Required. The ProgID or the CLSID of the object to create.
'   - A ProgID such as "MSCAL.Calendar.7"
'   - A CLSID such as "{8E27C92B-1264-101C-8A2F-040224009C02}"
' Return value: The dispatch interface of the control.
' ========================================================================================
PRIVATE FUNCTION CAxHost.CreateObject (BYREF wszProgID AS WSTRING) AS IDispatch PTR

   CAXH_DP("CAxHist.CreateObject - ProgID: " & wszProgID)
   ' // Exit if wszProgID is a null string
   IF LEN(wszProgID) = 0 THEN RETURN NULL
   ' // Convert the ProgID in a CLSID
   DIM ClassID AS CLSID
   CLSIDFromProgID(@wszProgID, @ClassID)
   ' // If it fails, see if it is a CLSID
   IF IsEqualIID(@ClassID, @IID_NULL) THEN
      ClassID = this.GuidFromStr(@wszProgID)
   END IF
'   ' // If not a valid ProgID or CLSID return NULL
   IF IsEqualIID(@ClassID, @IID_NULL) THEN RETURN NULL

   ' // Create an instance of the object
   DIM pDisp AS IDispatch PTR
   DIM hr AS HRESULT = CoCreateInstance(@ClassID, NULL, CLSCTX_SERVER, @IID_IDispatch, cast(LPVOID, @pDisp))
   CAXH_DP("CAxHost.CreateObject - hr - " & STR(hr) & " - IDispatch = " & WSTR(pDisp))
   RETURN pDisp

END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates an instance of an object.
' Parameters:
' - wszProgID: Required. The ProgID or the CLSID of the object to create.
'   - A ProgID such as "MSCAL.Calendar.7"
'   - A CLSID such as "{8E27C92B-1264-101C-8A2F-040224009C02}"
' - wszLicKey: The license key of the control.
' Return value: The dispatch interface of the control.
' ========================================================================================
PRIVATE FUNCTION CAxHost.CreateObject (BYREF wszProgID AS WSTRING, BYREF wszLicKey AS WSTRING) AS IDispatch PTR

   CAXH_DP("CAxHost.CreateObject - ProgID: " & wszProgID)
   ' // Exit if wszProgID is a null string
   IF LEN(wszProgID) = 0 THEN RETURN NULL
   ' // Convert the ProgID in a CLSID
   DIM ClassID  AS CLSID
   CLSIDFromProgID(@wszProgID, @ClassID)
   ' // If it fails, see if it is a CLSID
   IF IsEqualIID(@ClassID, @IID_NULL) THEN ClassID = this.GuidFromStr(@wszProgID)
'   ' // If not a valid CLSID return NULL
   IF IsEqualIID(@ClassID, @IID_NULL) THEN RETURN NULL

   ' // Get a reference to the IClassFactory2 interface of the control
   DIM pIClassFactory2 AS IClassFactory2 PTR, pDisp AS IDispatch PTR
   DIM hr AS HRESULT = CoGetClassObject(@ClassID, CLSCTX_SERVER, NULL, @IID_IClassFactory2, @pIClassFactory2)
   CAXH_DP("CoGetClassObject - hr - " & STR(hr) & " - IClassFactory2 = " & WSTR(pIClassFactory2))
   IF hr = S_OK AND pIClassFactory2 <> NULL THEN
      ' // Create a licensed instance of the control
      hr = pIClassFactory2->lpvtbl->CreateInstanceLic(pIClassFactory2, NULL, NULL, @IID_IDispatch, @wszLicKey, @pDisp)
      CAXH_DP("CreateInstanceLic - hr - " & STR(hr) & " - IDispatch = " & WSTR(pDisp))
      RETURN pDisp
   END IF

END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates an instance of an object.
' Parameters:
' - classID: A CLSID such as "{8E27C92B-1264-101C-8A2F-040224009C02}"
' - riid: An IID such "{C1AFAFB8-03E9-4E18-8DA5-51FBFF27CC66}"
' Return value: The dispatch interface of the control.
' ========================================================================================
PRIVATE FUNCTION CAxHost.CreateObject (BYREF classID AS CONST CLSID, BYREF riid AS CONST IID) AS IDispatch PTR
   CAXH_DP("CAxHost.CreateObject - classID & riid")
   ' // Create an instance of the object
   DIM pDisp AS IDIspatch PTR
   DIM hr AS HRESULT = CoCreateInstance(@ClassID, NULL, CLSCTX_SERVER, @riid, cast(LPVOID, @pDisp))
   CAXH_DP("CAxHost.CreateObject - hr - " & STR(hr) & " - IDIspatch = " & WSTR(pDisp))
   RETURN pDisp
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates an instance of an object.
' Parameters:
' - classID: A CLSID such as "{8E27C92B-1264-101C-8A2F-040224009C02}"
' - riid: An IID such "{C1AFAFB8-03E9-4E18-8DA5-51FBFF27CC66}"
' - wszLicKey: The license key of the control.
' Return value: The dispatch interface of the control.
' ========================================================================================
PRIVATE FUNCTION CAxHost.CreateObject (BYREF classID AS CONST CLSID, BYREF riid AS CONST IID, BYREF wszLicKey AS WSTRING) AS IDispatch PTR
   CAXH_DP("CAxHost.CreateObject - classID & riid & LicKey")
   ' // Get a reference to the IClassFactory2 interface of the control
   DIM pIClassFactory2 AS IClassFactory2 PTR, pDisp AS IDispatch PTR
   DIM hr AS HRESULT = CoGetClassObject(@classID, CLSCTX_SERVER, NULL, @IID_IClassFactory2, @pIClassFactory2)
   CAXH_DP("CoGetClassObject - hr - " & STR(hr) & " - IClassFactory2 = " & WSTR(pIClassFactory2))
   IF hr = S_OK AND pIClassFactory2 <> NULL THEN
      ' // Create a licensed instance of the control
      hr = pIClassFactory2->lpvtbl->CreateInstanceLic(pIClassFactory2, NULL, NULL, @riid, @wszLicKey, @pDisp)
      CAXH_DP("CreateInstanceLic - hr - " & STR(hr) & " - IDispatch = " & WSTR(pDisp))
      RETURN pDisp
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Loads the specified library from file and creates an instance of an object.
' Parameters:
' - wszLibName = Full path where the library is located.
' - rclsid = The CLSID (class identifier) associated with the data and code that will be
'   used to create the object.
' - riid = A reference to the identifier of the interface to be used to communicate with the object.
' - wszLicKey = The license key.
' If it succeeds, returns a reference to the requested interface; otherwise, it returns null.
' Not every component is a suitable candidate for use under this overloaded AfxNewCom function.
'  - Only in-process servers (DLLs) are supported.
'  - Components that are system components or part of the operating system, such as XML,
'    Data Access, Internet Explorer, or DirectX, aren't supported
'  - Components that are part of an application, such Microsoft Office, aren't supported.
'  - Components intended for use as an add-in or a snap-in, such as an Office add-in or
'    a control in a Web browser, aren't supported.
'  - Components that manage a shared physical or virtual system resource aren't supported.
'  - Visual ActiveX controls aren't supported because they need to be initilized and
'    activated by the OLE container.
' Note: Do not use DyLibFree to unload the library once you have got a valid reference
' to an interface or your application will GPF. Before calling DyLibFree, all the
' interface references must be released. If you don't need to unload the library until
' the application ends, then you don't need to call FreeLibrary because CoUninitialize
' closes the COM library on the current thread, unloads all DLLs loaded by the thread,
' frees any other resources that the thread maintains, and forces all RPC connections on
' the thread to close.
' ========================================================================================
PRIVATE FUNCTION CAxHost.CreateObject (BYREF wszLibName AS CONST WSTRING, BYREF rclsid AS CONST CLSID, _
   BYREF riid AS CONST IID, BYREF wszLicKey AS WSTRING = "") AS IDispatch PTR
   CAXH_DP("CAxHost.CreateObject - LibName & classID & riid & LicKey")
   CAXH_DP("CAxHost.CreateObject - LibName = " & wszLibName)
   CAXH_DP("CAxHost.CreateObject - LicKey = " & wszLicKey)

   ' // See if the library is already loaded in the address space
   DIM hLib AS HMODULE = GetModuleHandleW(wszLibName)
   ' // If it is not loaded, load it
   IF hLib = NULL THEN hLib = LoadLibraryW(wszLibName)
   CAXH_DP("CAxHost.CreateObject - hLib = " & WSTR(hLib))
   ' // If it fails, abort
   IF hLib = NULL THEN RETURN NULL

   ' // Retrieve the address of the exported function DllGetClassObject
   DIM pfnDllGetClassObject AS FUNCTION (BYVAL rclsid AS CONST IID CONST PTR, BYVAL riid AS CONST IID CONST PTR, BYVAL ppv AS LPVOID PTR) AS HRESULT
   pfnDllGetClassObject = CAST(ANY PTR, GetProcAddress(hLib, "DllGetClassObject"))
   CAXH_DP("CAxHost.CreateObject - DllGetClassObject = " & WSTR(pfnDllGetClassObject))
   IF pfnDllGetClassObject = NULL THEN RETURN NULL

   IF LEN(wszLicKey) = 0 THEN
      ' // Request a reference to the IClassFactory interface
      DIM pIClassFactory AS IClassFactory PTR, pDisp AS IDispatch PTR
      DIM hr AS HRESULT = pfnDllGetClassObject(@rclsid, @IID_IClassFactory, @pIClassFactory)
      CAXH_DP("CAxHost.CreateObject - IClassFactory = " & WSTR(pIClassFactory))
      IF hr <> S_OK THEN RETURN NULL
      ' // Create an instance of the server or control
      hr = pIClassFactory->lpVtbl->CreateInstance(pIClassFactory, NULL, @riid, @pDisp)
      pIClassFactory->lpVtbl->Release(pIClassFactory)
      IF hr <> S_OK THEN RETURN NULL
      RETURN pDisp
   ELSE
      ' // Request a reference to the IClassFactory2 interface
      DIM pIClassFactory2 AS IClassFactory2 PTR, pDisp AS IDispatch PTR
      DIM hr AS HRESULT = pfnDllGetClassObject(@rclsid, @IID_IClassFactory, @pIClassFactory2)
      CAXH_DP("CAxHost.CreateObject - IClassFactory2 = " & WSTR(pIClassFactory2))
      IF hr <> S_OK THEN RETURN NULL
      ' // Create a licensed instance of the server or control
      hr = pIClassFactory2->lpVtbl->CreateInstanceLic(pIClassFactory2, NULL, NULL, @riid, @wszLicKey, @pDisp)
      CAXH_DP("CAxHost.CreateObject - IClassFactory2 - hr = &h" & HEX(hr))
      pIClassFactory2->lpVtbl->Release(pIClassFactory2)
      IF hr <> S_OK THEN RETURN NULL
      RETURN pDisp
   END IF

END FUNCTION
' ========================================================================================

' ========================================================================================
' OLE container control window procedure
' ========================================================================================
PRIVATE FUNCTION CAxHost.CAxHostProc (BYVAL hWnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   DIM hr AS HRESULT

   SELECT CASE uMsg

      ' ==================================================================================
      CASE WM_CREATE
      ' ==================================================================================
         CAXH_DP("--- WM_CREATE BEGIN ---")

         ' ===============================================================================
         ' // Get a pointer to the OLE Container class
         DIM pCreateStruct AS CREATESTRUCT PTR = CAST(CREATESTRUCT PTR, lParam)
         DIM pAxHost AS CAxHost PTR = CAST(CAxHost PTR, pCreateStruct->lpCreateParams)
         IF pAxHost = NULL THEN RETURN -1
         ' // Store the handle of the window
         pAxHost->m_hwndContainer = hwnd
         ' ===============================================================================

         ' ===============================================================================
         ' // Create an instance of the ActiveX control
         IF LEN(pAxHost->m_wszLibName) THEN
            pAxHost->m_pOcx = pAxHost->CreateObject(pAxHost->m_wszLibName, _
               pAxHost->m_rclsid, pAxHost->m_riid, pAxHost->m_wszLicKey)
         ELSEIF LEN(pAxHost->m_wszProgID) <> 0 AND LEN(pAxHost->m_wszLicKey) <> 0 THEN
            pAxHost->m_pOcx = pAxHost->CreateObject(pAxHost->m_wszProgID, pAxHost->m_wszLicKey)
         ELSEIF LEN(pAxHost->m_wszProgID) <> 0 THEN
            pAxHost->m_pOcx = pAxHost->CreateObject(pAxHost->m_wszProgID)
         ELSEIF IsEqualIID(@pAxHost->m_rclsid, @IID_NULL) = FALSE AND IsEqualIID(@pAxHost->m_riid, @IID_NULL) = FALSE AND LEN(pAxHost->m_wszLicKey) <> 0 THEN
            pAxHost->m_pOcx = pAxHost->CreateObject(pAxHost->m_rclsid, pAxHost->m_riid, pAxHost->m_wszLicKey)
         ELSEIF IsEqualIID(@pAxHost->m_rclsid, @IID_NULL) = FALSE AND IsEqualIID(@pAxHost->m_riid, @IID_NULL) = FALSE THEN
            pAxHost->m_pOcx = pAxHost->CreateObject(pAxHost->m_rclsid, pAxHost->m_riid)
         END IF
        ' // If it fails, abort the creation of the control
         IF pAxHost->m_pOcx = NULL THEN RETURN -1
         ' ===============================================================================

         ' ===============================================================================
         ' // Creates a default font
         IF pAxHost->m_AmbientDisp.Font = NULL THEN
            IF AfxWindowsVersion >= 600 AND CLNG(AfxIsProcessDpiAware) THEN
               pAxHost->m_AmbientDisp.Font = pAxHost->OleCreateFontDisp("Segoe UI", 9, FW_NORMAL, 0, 0, 0, DEFAULT_CHARSET)
            ELSE
               pAxHost->m_AmbientDisp.Font = pAxHost->OleCreateFontDisp("Tahoma", 8, FW_NORMAL, 0, 0, 0, DEFAULT_CHARSET)
            END IF
         END IF
         ' ===============================================================================

         ' ===============================================================================
         ' // Retrieves the IOleObject interface
         pAxHost->m_pOcx->lpvtbl->QueryInterface(pAxHost->m_pOcx, @IID_IOleObject, @pAxHost->m_pIOleObject)
         ' // If it doesn't exist, it's not a visual control
         IF pAxHost->m_pIOleObject = NULL THEN RETURN -1
         ' ===============================================================================

         ' ===============================================================================
         ' // Gets the status bitwise flags
         pAxHost->m_pIOleObject->GetMiscStatus(DVASPECT_CONTENT, @pAxHost->m_dwMiscStatus)
         ' // The control needs to be put on running state
         IF (pAxHost->m_dwMiscStatus AND OLEMISC_ALWAYSRUN) = OLEMISC_ALWAYSRUN THEN
            hr = OleRun(cast(IUnknown PTR, pAxHost->m_pOcx))
            CAXH_DP("OleRun: hr = " & HEX(hr))
         END IF
         IF (pAxHost->m_dwMiscStatus AND OLEMISC_SIMPLEFRAME) = OLEMISC_SIMPLEFRAME THEN
            CAXH_DP("--- Simple frame ---")
         END IF
         ' // The control wants that you set the client site first
         IF (pAxHost->m_dwMiscStatus AND OLEMISC_SETCLIENTSITEFIRST) = OLEMISC_SETCLIENTSITEFIRST THEN
            ' // Sets the client site
            pAxHost->QueryInterface(@IID_IOleClientSite, @pAxHost->m_pIOleClientSiteImpl)
            IF pAxHost->m_pIOleClientSiteImpl THEN
               CAXH_DP("--- SetClientSite - Begin ---")
               hr = pAxHost->m_pIOleObject->SetClientSite(CAST(ANY PTR, pAxHost->m_pIOleClientSiteImpl))
               CAXH_DP("--- SetClientSite - End --- hr = " & HEX(hr))
            END IF
         END IF
         ' ===============================================================================

         ' ===============================================================================
         ' // Sets the host name
         DIM wszContainerApp AS WSTRING * 260 = CAXHOST_CLASSNAME
         hr = pAxHost->m_pIOleObject->SetHostNames(@wszContainerApp, NULL)
         CAXH_DP("SetHostNames: " & HEX(hr))
         ' ===============================================================================

         ' ===============================================================================
         ' // Initializes the object to a default state
         DIM psi AS IPersistStreamInit PTR
         hr = IUnknown_QueryInterface(pAxHost->m_pOcx, @IID_IPersistStreamInit, @psi)
         IF hr = S_OK AND psi <> NULL THEN
            hr = psi->lpvtbl->InitNew(psi)
            CAXH_DP("IPersistStreamInit.InitNew - hr = " & HEX(hr))
            AfxSafeRelease(psi)
         END IF
         IF (pAxHost->m_dwMiscStatus AND OLEMISC_SETCLIENTSITEFIRST) <> OLEMISC_SETCLIENTSITEFIRST THEN
            ' // Sets the client site
            pAxHost->QueryInterface(@IID_IOleClientSite, @pAxHost->m_pIOleClientSiteImpl)
            IF pAxHost->m_pIOleClientSiteImpl THEN
               CAXH_DP("--- SetClientSite - Begin ---")
               hr = pAxHost->m_pIOleObject->SetClientSite(CAST(ANY PTR, pAxHost->m_pIOleClientSiteImpl))
               CAXH_DP("--- SetClientSite - End --- hr = " & HEX(hr))
            END IF
         END IF
         ' ===============================================================================

         ' ===============================================================================
         ' // Gets the size of the container's window
         DIM rc AS RECT
         GetClientRect pAxHost->m_hwndContainer, @rc
         ' // Converts from pixel to himetric
         DIM pxSize AS SIZEL, hmSize AS SIZEL
         pxSize.cx = rc.Right - rc.Left
         pxSize.cy = rc.Bottom - rc.Top
         pAxHost->PixelsToHiMetric(@pxSize, @hmSize)
         ' ===============================================================================

         ' ===============================================================================
         ' // Activates the control in place
         IF pAxHost->m_pIOleClientSiteImpl THEN
            CAXH_DP("--- DoVerb - INPLACEACTIVATE - Begin ---")
            DIM pIOleClientSite AS ANY PTR = pAxHost->m_pIOleClientSiteImpl
            DIM hr AS HRESULT = pAxHost->m_pIOleObject->DoVerb(OLEIVERB_INPLACEACTIVATE, NULL, _
               cast(IOleClientSite PTR, pIOleClientSite), 0, pAxHost->m_hwndContainer, @rc)
            CAXH_DP("--- DoVerb - INPLACEACTIVATE - End --- hr = " & HEX(hr))
            ' // Sets the extent of object's display area
            IF pAxHost->m_pIOleObject THEN pAxHost->m_pIOleObject->SetExtent(DVASPECT_CONTENT, @hmSize)
            ' // Indicates how much of the in-place object is visible.
            IF pAxHost->m_pIOleInPlaceObject = NULL THEN
               IUnknown_QueryInterface(pAxHost->m_pOcx, @IID_IOleInPlaceObject, @pAxHost->m_pIOleInPlaceObject)
            END IF
            IF pAxHost->m_pIOleInPlaceObject THEN
               pAxHost->m_pIOleInPlaceObject->SetObjectRects(@rc, @rc)
            END IF
         END IF
         ' ===============================================================================

         CAXH_DP("--- WM_CREATE END ---")
         EXIT FUNCTION

      ' ==================================================================================
      CASE WM_ERASEBKGND
      ' ==================================================================================
         ' // Don't erase the background to avoid flicker
         FUNCTION = 1
         EXIT FUNCTION

      ' ==================================================================================
   	CASE WM_ENABLE
      ' ==================================================================================
         InvalidateRect hwnd, NULL, 0
         UpdateWindow hwnd
         EXIT FUNCTION

      ' ==================================================================================
      CASE WM_SETFOCUS
      ' ==================================================================================
         CAXH_DP("--- WM_SETFOCUS BEGIN ---")
         DIM pAxHost AS CAxHost PTR = CAST(CAxHost PTR, GetWindowLongPtrW(hwnd, 0))
         IF pAxHost = NULL THEN EXIT FUNCTION
         pAxHost->m_bHaveFocus = TRUE
         ' // Activates an object in place, along with its full set of user-interface tools.
         IF pAxHost->m_pIOleObject <> NULL AND CLNG(pAxHost->m_bInPlaceActive) = FALSE THEN
            DIM rc AS RECT
            GetClientRect hwnd, @rc
            IF pAxHost->m_pIOleClientSiteImpl THEN
               DIM pIOleClientSite AS ANY PTR = pAxHost->m_pIOleClientSiteImpl
               pAxHost->m_pIOleObject->DoVerb(OLEIVERB_UIACTIVATE, NULL, pIOleClientSite, 0, hwnd, @rc)
               CAXH_DP("--- WM_SETFOCUS OLEIVERB_UIACTIVATE ---")
            END IF
         END IF
         IF pAxHost->m_pIOleInPlaceObject = NULL THEN
            IUnknown_QueryInterface(pAxHost->m_pOcx, @IID_IOleInPlaceObject, @pAxHost->m_pIOleInPlaceObject)
         END IF
         IF pAxHost->m_pIOleInPlaceObject THEN
            DIM hwndCtrl AS HANDLE
            pAxHost->m_pIOleInPlaceObject->GetWindow(@hwndCtrl)
            IF IsWindow(hwndCtrl) THEN SetFocus hwndCtrl
            CAXH_DP("--- WM_SETFOCUS hwndCtrl = " & WSTR(hwndCtrl) & " ---")
         END IF
         CAXH_DP("--- WM_SETFOCUS END ---")
         EXIT FUNCTION

      ' ==================================================================================
      CASE WM_KILLFOCUS
      ' ==================================================================================
         CAXH_DP("--- WM_KILLFOCUS BEGIN ---")
         DIM pAxHost AS CAxHost PTR = CAST(CAxHost PTR, GetWindowLongPtrW(hwnd, 0))
         IF pAxHost = NULL THEN EXIT FUNCTION
         pAxHost->m_bHaveFocus = FALSE
         CAXH_DP("--- WM_KILLFOCUS END ---")
         EXIT FUNCTION

      ' ==================================================================================
      ' Resizes the control to fill the client area of the host window
      ' ==================================================================================
      CASE WM_SIZE
      ' ==================================================================================
         DIM pAxHost AS CAxHost PTR = CAST(CAxHost PTR, GetWindowLongPtrW(hwnd, 0))
         IF pAxHost = NULL THEN EXIT FUNCTION
         CAXH_DP("--- WM_SIZE BEGIN ---")
         IF pAxHost->m_bInPlaceActive THEN
            ' // Retrieves the coordinates of a window's client area.
            DIM rc AS RECT
            GetClientRect hwnd, @rc
            ' // Converts from pixel to himetric
            DIM pxSize AS SIZEL, hmSize AS SIZEL
            pxSize.cx = rc.Right - rc.Left
            pxSize.cy = rc.Bottom - rc.Top
            pAxHost->PixelsToHiMetric(@pxSize, @hmSize)
            ' // Sets extent of object's display area
            IF pAxHost->m_pIOleObject THEN pAxHost->m_pIOleObject->SetExtent(DVASPECT_CONTENT, @hmSize)
            ' // Indicates how much of the in-place object is visible.
            IF pAxHost->m_pIOleInPlaceObject = NULL THEN
               IUnknown_QueryInterface(pAxHost->m_pOcx, @IID_IOleInPlaceObject, @pAxHost->m_pIOleInPlaceObject)
            END IF
            IF pAxHost->m_pIOleInPlaceObject THEN
               pAxHost->m_pIOleInPlaceObject->SetObjectRects(@rc, @rc)
            END IF
         END IF
         CAXH_DP("--- WM_SIZE END ---")
         EXIT FUNCTION

   END SELECT

   ' // Default processing for other messages.
   FUNCTION = DefWindowProcW(hWnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the interface name for debugging purposes
' ========================================================================================
PRIVATE FUNCTION CAxHost.GetInterfaceName (BYVAL riid AS REFIID) AS STRING

   IF IsEqualIID(riid, @IID_IUnknown) THEN
      FUNCTION = "IID_IUnknown"
   ELSEIF IsEqualIID(riid, @IID_IDispatch) THEN
      FUNCTION = "IID_IDispatch"
   ELSEIF IsEqualIID(riid, @IID_IAdviseSink) THEN
      FUNCTION = "IID_IAdviseSink"
   ELSEIF IsEqualIID(riid, @IID_IErrorInfo) THEN
      FUNCTION = "IID_IErrorInfo"
   ELSEIF IsEqualIID(riid, @IID_IOleClientSite) THEN
      FUNCTION = "IID_IOleClientSite"
   ELSEIF IsEqualIID(riid, @IID_IOleContainer) THEN
      FUNCTION = "IID_IOleContainer"
   ELSEIF IsEqualIID(riid, @IID_IOleControl) THEN
      FUNCTION = "IID_IOleControl"
   ELSEIF IsEqualIID(riid, @IID_IOleControlSite) THEN
      FUNCTION = "IID_IOleControlSite"
   ELSEIF IsEqualIID(riid, @IID_IOleDocumentSite) THEN
      FUNCTION = "IID_IOleDocumentSite"
   ELSEIF IsEqualIID(riid, @IID_IOleInPlaceActiveObject) THEN
      FUNCTION = "IID_IOleInPlaceActiveObject"
   ELSEIF IsEqualIID(riid, @IID_IOleInPlaceFrame) THEN
      FUNCTION = "IID_IOleInPlaceFrame"
   ELSEIF IsEqualIID(riid, @IID_IOleInPlaceObject) THEN
      FUNCTION = "IID_IOleInPlaceObject"
   ELSEIF IsEqualIID(riid, @IID_IOleInPlaceObjectWindowless) THEN
      FUNCTION = "IID_IOleInPlaceObjectWindowless"
   ELSEIF IsEqualIID(riid, @IID_IOleInPlaceSite) THEN
      FUNCTION = "IID_IOleInPlaceSite"
   ELSEIF IsEqualIID(riid, @IID_IOleInPlaceSiteEx) THEN
      FUNCTION = "IID_IOleInPlaceSiteEx"
   ELSEIF IsEqualIID(riid, @IID_IOleInPlaceUIWindow) THEN
      FUNCTION = "IID_IOleInPlaceUIWindow"
   ELSEIF IsEqualIID(riid, @IID_IOleObject) THEN
      FUNCTION = "IID_IOleObject"
   ELSEIF IsEqualIID(riid, @IID_IOleWindow) THEN
      FUNCTION = "IID_IOleWindow"
   ELSEIF IsEqualIID(riid, @IID_IParseDisplayName) THEN
      FUNCTION = "IID_IParseDisplayName"
   ELSEIF IsEqualIID(riid, @IID_IPersistStreamInit) THEN
      FUNCTION = "IID_IPersistStreamInit"
   ELSEIF IsEqualIID(riid, @IID_IPropertyNotifySink) THEN
      FUNCTION = "IID_IPropertyNotifySink"
   ELSEIF IsEqualIID(riid, @IID_IServiceProvider) THEN
      FUNCTION = "IID_IServiceProvider"
   ELSEIF IsEqualIID(riid, @IID_ISimpleFrameSite) THEN
      FUNCTION = "IID_ISimpleFrameSite"
   ' // VB-only interfaces - Listed here for debugging purposes
'   ELSEIF IsEqualIID(riid, @IID_IVBGetControl) THEN
'      FUNCTION = "IID_IVBGetControl"
'   ELSEIF IsEqualIID(riid, @IID_IGetOleObject) THEN
'      FUNCTION = "IID_IGetOleObject"
'   ELSEIF IsEqualIID(riid, @IID_IVBFormat) THEN
'      FUNCTION = "IID_IVBFormat"
'   ELSEIF IsEqualIID(riid, @IID_IGetVBAObject) THEN
'      FUNCTION = "IID_IGetVBAObject"
   ELSE
      FUNCTION = this.GuidText(riid)
   END IF

END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
FUNCTION CAxHost.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   CAXH_DP("CAxHost.QueryInterface : " & this.GetInterfaceName(riid))

   IF IsEqualIID(riid, @IID_IUnknown) OR IsEqualIID(riid, @IID_IDispatch) THEN
      IF this.m_pIDispatchImpl THEN
         this.m_pIDispatchImpl->AddRef
         *ppvObj = this.m_pIDispatchImpl
      ELSE
         DIM p AS ANY PTR = NEW CAxHost_IDispatch(@this)
         this.m_pIDispatchImpl = p
         this.m_pIDispatchImpl->AddRef
         *ppvObj = p
      END IF
   ELSEIF IsEqualIID(riid, @IID_IOleClientSite) THEN
      IF this.m_pIOleClientSiteImpl THEN
         this.m_pIOleClientSiteImpl->AddRef
         *ppvObj = this.m_pIOleClientSiteImpl
      ELSE
         DIM p AS ANY PTR = NEW CAxHost_IOleClientSite(@this)
         this.m_pIOleClientSiteImpl = p
         this.m_pIOleClientSiteImpl->AddRef
         *ppvObj = p
      END IF
   ELSEIF IsEqualIID(riid, @IID_IOleControlSite) THEN
      IF this.m_pIOleControlSiteImpl THEN
         this.m_pIOleControlSiteImpl->AddRef
         *ppvObj = this.m_pIOleControlSiteImpl
      ELSE
         DIM p AS ANY PTR = NEW CAxHost_IOleControlSite(@this)
         this.m_pIOleControlSiteImpl = p
         this.m_pIOleControlSiteImpl->AddRef
         *ppvObj = p
      END IF
   ELSEIF IsEqualIID(riid, @IID_IOleInPlaceFrame) OR IsEqualIID(riid, @IID_IOleWindow) OR IsEqualIID(riid, @IID_IOleInPlaceUIWindow) THEN
      IF this.m_pIOleInPlaceFrameImpl THEN
         this.m_pIOleInPlaceFrameImpl->AddRef
         *ppvObj = this.m_pIOleInPlaceFrameImpl
      ELSE
         DIM p AS ANY PTR = NEW CAxHost_IOleInPlaceFrame(@this)
         this.m_pIOleInPlaceFrameImpl = p
         this.m_pIOleInPlaceFrameImpl->AddRef
         *ppvObj = p
      END IF
   ELSEIF IsEqualIID(riid, @IID_IOleInPlaceSite) THEN
      IF this.m_pIOleInPlaceSiteImpl THEN
         this.m_pIOleInPlaceSiteImpl->AddRef
         *ppvObj = this.m_pIOleInPlaceSiteImpl
      ELSE
         DIM p AS ANY PTR = NEW CAxHost_IOleInPlaceSite(@this)
         this.m_pIOleInPlaceSiteImpl = p
         this.m_pIOleInPlaceSiteImpl->AddRef
         *ppvObj = p
      END IF
   ELSEIF IsEqualIID(riid, @IID_IServiceProvider) THEN
      IF this.m_pIServiceProviderImpl THEN
         this.m_pIServiceProviderImpl->AddRef
         *ppvObj = this.m_pIServiceProviderImpl
      ELSE
         DIM p AS ANY PTR = NEW CAxHost_IServiceProvider(@this)
         this.m_pIServiceProviderImpl = p
         this.m_pIServiceProviderImpl->AddRef
         *ppvObj = p
      END IF
   END IF

   IF *ppvObj = NULL THEN RETURN E_NOINTERFACE
   RETURN S_OK

END FUNCTION
' ========================================================================================

END NAMESPACE

' ########################################################################################
'                              *** HELPER PROCEDURES ***
' ########################################################################################

USING Afx

' ========================================================================================
' Returns the OLE container window handle given the handle of the form or any control in the
' form and the control identifier.
' Parameters:
' - hwnd = Reference window handle.
' - Control identifier, e.g. IDC_WEBBROWSER.
' Returns the OLE container window handle or NULL.
' ========================================================================================
PRIVATE FUNCTION AfxCAxHostWindow (BYVAL hwnd AS HWND, BYVAL cID AS WORD) AS HWND
   DIM wszClassName AS WSTRING * 260
   DIM hwndChild AS .HWND = AfxGetControlHandle(hwnd, cID)
   GetClassNameW hwndChild, wszClassName, SIZEOF(wszClassName)
   IF wszClassName <> CAXHOST_CLASSNAME THEN RETURN NULL
   RETURN hwndChild
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a pointer to the CAxHost class given the handle of its associated window.
' ========================================================================================
PRIVATE FUNCTION AfxCAxHostPtr OVERLOAD (BYVAL hwnd AS HWND) AS CAxHost PTR
   DIM wszClassName AS WSTRING * 260
   GetClassNameW hwnd, wszClassName, SIZEOF(wszClassName)
   IF wszClassName <> CAXHOST_CLASSNAME THEN RETURN NULL
   RETURN CAST(CAxHost PTR, .GetWindowLongPtrW(hwnd, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a pointer to the CAxHost class given the handle of the form or any control in the
' form and the control identifier.
' Parameters:
' - hwnd = Reference window handle.
' - Control identifier, e.g. IDC_WEBBROWSER.
' Returns a pointer to the COleCon class or NULL.
' ========================================================================================
PRIVATE FUNCTION AfxCAxHostPtr OVERLOAD (BYVAL hwnd AS HWND, BYVAL cID AS WORD) AS CAxHost PTR
   DIM wszClassName AS WSTRING * 260
   DIM hwndChild AS .HWND = AfxGetControlHandle(hwnd, cID)
   GetClassNameW hwndChild, wszClassName, SIZEOF(wszClassName)
   IF wszClassName <> CAXHOST_CLASSNAME THEN RETURN NULL
   RETURN CAST(CAxHost PTR, .GetWindowLongPtrW(hwndChild, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a reference to the hosted ActiveX control dispatch interface given the handle of
' the window that hosts it. Don't call IUnknown_Release.
' ========================================================================================
PRIVATE FUNCTION AfxCAxHostDispPtr (BYVAL hwnd AS HWND) AS IDispatch PTR
   DIM wszClassName AS WSTRING * 260
   GetClassNameW hwnd, wszClassName, SIZEOF(wszClassName)
   IF wszClassName <> CAXHOST_CLASSNAME THEN RETURN NULL
   DIM pAxHost AS CAxHost PTR = CAST(CAxHost PTR, .GetWindowLongPtrW(hwnd, 0))
   IF pAxHost THEN RETURN pAxHost->m_pOcx
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns a counted reference to the hosted ActiveX control dispatch interface given the
' handle of the window that hosts it. You must call IUnknown_Release when no longer need it.
' ========================================================================================
PRIVATE FUNCTION AfxCAxHostDispObj (BYVAL hwnd AS HWND) AS IDispatch PTR
   DIM wszClassName AS WSTRING * 260
   GetClassNameW hwnd, wszClassName, SIZEOF(wszClassName)
   IF wszClassName <> CAXHOST_CLASSNAME THEN RETURN NULL
   DIM pAxHost AS CAxHost PTR = CAST(CAxHost PTR, .GetWindowLongPtrW(hwnd, 0))
   IF pAxHost THEN
      IF pAxHost->m_pOcx <> NULL THEN
         pAxHost->m_pOcx->lpvtbl->AddRef(pAxHost->m_pOcx)
         RETURN pAxHost->m_pOcx
      END IF
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Forwards the message to the control. Active in-place objects must always be given the
' first chance at translating accelerator keystrokes. You can provide this opportunity by
' calling IOleInPlaceActiveObject.TranslateAccelerator from your container's message loop
' before doing any other translation. You should apply your own translation only when this
' method returns FALSE.
' Remarks: If the control is the WebBrowser control, TranslateAccelerator calls the namesake
' method of the IDocHostUIHandler interface.
' Usage example:
'   ' // Dispatch Windows messages
'   DIM uMsg AS MSG
'   WHILE (GetMessageW(@uMsg, NULL, 0, 0) <> FALSE)
'      IF AfxCAxHostForwardMessage(GetFocus, @uMsg) = FALSE THEN
'         IF IsDialogMessageW(hWndMain, @uMsg) = 0 THEN
'            TranslateMessage(@uMsg)
'            DispatchMessageW(@uMsg)
'         END IF
'      END IF
'   WEND
'   FUNCTION = uMsg.wParam
' ========================================================================================
PRIVATE FUNCTION AfxCAxHostForwardMessage (BYVAL hctl AS HWND, BYVAL pMsg AS tagMsg PTR) AS BOOLEAN
   IF pMsg->message >= WM_KEYFIRST AND pMsg->message <= WM_KEYLAST THEN
      ' // See if the window that has the focus is a child of our container window
      DIM wszClassName AS WSTRING * 260
      DIM hWndParent AS HWND, hWndTmp AS HWND
      hWndParent = hCtl
      DO
         hWndTmp = GetParent(hWndParent)
         IF hWndTmp = NULL THEN EXIT DO
         hWndParent = hWndTmp
         GetClassNameW hWndParent, wszClassName, SIZEOF(wszClassName)
         IF wszClassName = CAXHOST_CLASSNAME THEN EXIT DO
      LOOP
      IF wszClassName <> CAXHOST_CLASSNAME THEN RETURN FALSE
      ' // Get a pointer to the OLE container class
      DIM pAxHost AS CAxHost PTR = cast(CAxHost PTR, GetWindowLongPtrW(hwndParent, 0))
      IF pAxHost = NULL THEN RETURN FALSE
      IF pAxHost->m_pOcx = NULL THEN RETURN FALSE
      ' // Translate the message
      DIM hr AS HRESULT, pActiveObject AS IOleInPlaceActiveObject PTR
      hr = IUnknown_QueryInterface(pAxHost->m_pOcx, @IID_IOleInPlaceActiveObject, @pActiveObject)
      IF pActiveObject = NULL THEN RETURN FALSE
      hr = pActiveObject->lpvtbl->TranslateAccelerator(pActiveObject, pMsg)
      pActiveObject->lpvtbl->Release(pActiveObject)
      IF hr = S_OK THEN RETURN TRUE
   END IF
   RETURN FALSE
END FUNCTION
' ========================================================================================
