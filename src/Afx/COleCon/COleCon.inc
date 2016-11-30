' ########################################################################################
' Microsoft Windows
' File: COleCon.inc
' Contents: OLE Container class
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once

TYPE Afx_IOleClientSiteImpl_ AS Afx_IOleClientSiteImpl
TYPE Afx_IServiceProviderImpl_ AS Afx_IServiceProviderImpl
TYPE Afx_IParseDisplayNameImpl_ AS Afx_IParseDisplayNameImpl
TYPE Afx_IOleContainerImpl_ AS Afx_IOleContainerImpl
TYPE Afx_IOleWindowImpl_ AS Afx_IOleWindowImpl
TYPE Afx_IOleInPlaceUIWindowImpl_ AS Afx_IOleInPlaceUIWindowImpl
TYPE Afx_IOleInPlaceFrameImpl_ AS Afx_IOleInPlaceFrameImpl
TYPE Afx_IOleInPlaceSiteImpl_ AS Afx_IOleInPlaceSiteImpl
TYPE Afx_IOleInPlaceSiteExImpl_ AS Afx_IOleInPlaceSiteExImpl
TYPE Afx_IOleControlSiteImpl_ AS Afx_IOleControlSiteImpl
TYPE Afx_ISimpleFrameSiteImpl_ AS Afx_ISimpleFrameSiteImpl

#include once "Afx/COleCon/COleCon.bi"
#include once "Afx/COleCon/COleConDispatch.inc"

' ========================================================================================
' COleCon class
' Implements an OLE container as a Windows custom control.
' ========================================================================================
TYPE COleCon

Public:
   m_pData AS Afx_OleConDispatch PTR
   m_bUninitOLE AS BOOLEAN
   m_dwCookie AS DWORD          ' // Cookie for Unadvise
   m_riidEvt AS IID PTR         ' // IID of the events interface

Private:
   DECLARE SUB CreateContainer (BYVAL pWindow AS CWindow PTR, BYVAL cID AS INTEGER, BYREF wszProgID AS WSTRING, _
      BYREF wszLibName AS WSTRING, BYREF rclsid AS const CLSID, BYREF riid AS const IID, BYREF wszLicKey AS WSTRING, _
      BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
      BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, BYVAL ambientFlags AS DWORD = 0)

Public:
   DECLARE CONSTRUCTOR (BYVAL pWindow AS CWindow PTR, BYVAL cID AS INTEGER, BYREF wszLibName AS WSTRING, _
      BYREF rclsid AS const CLSID, BYREF riid AS const IID, BYREF wszLicKey AS WSTRING, _
      BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
      BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, BYVAL ambientFlags AS DWORD = 0)
   DECLARE CONSTRUCTOR (BYVAL pWindow AS CWindow PTR, BYVAL cID AS INTEGER, BYREF wszLibName AS WSTRING, _
      BYREF rclsid AS const CLSID, BYREF riid AS const IID, BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, _
      BYVAL nHeight AS LONG = 0, BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, BYVAL ambientFlags AS DWORD = 0)
   DECLARE CONSTRUCTOR (BYVAL pWindow AS CWindow PTR, BYVAL cID AS INTEGER, BYREF wszProgID AS WSTRING, _
      BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
      BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, BYVAL ambientFlags AS DWORD = 0)
   DECLARE CONSTRUCTOR (BYVAL pWindow AS CWindow PTR, BYVAL cID AS INTEGER, BYREF wszProgID AS WSTRING, _
      BYREF wszLicKey AS WSTRING, BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
      BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, BYVAL ambientFlags AS DWORD = 0)
   DECLARE CONSTRUCTOR (BYVAL pWindow AS CWindow PTR, BYVAL cID AS INTEGER, BYREF classID AS CONST CLSID, _
      BYREF riid AS CONST IID, BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, _
      BYVAL nHeight AS LONG = 0, BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, BYVAL ambientFlags AS DWORD = 0)
   DECLARE CONSTRUCTOR (BYVAL pWindow AS CWindow PTR, BYVAL cID AS INTEGER, BYREF classID AS CONST CLSID, _
      BYREF riid AS CONST IID, BYREF wszLicKey AS WSTRING, BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, _
      BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, BYVAL dwStyle AS DWORD = 0, _
      BYVAL dwExStyle AS DWORD = 0, BYVAL ambientFlags AS DWORD = 0)
   DECLARE DESTRUCTOR
   DECLARE FUNCTION RegisterClass () AS ATOM
   DECLARE FUNCTION OleCreateFontDisp (BYREF wszFontName AS WSTRING, BYVAL cySize AS LONGLONG, _
           BYVAL fWeight AS SHORT, BYVAL fCharset AS SHORT, BYVAL fItalic AS WINBOOL, _
           BYVAL fUnderline AS WINBOOL, BYVAL fStrikethrough AS WINBOOL) AS IFontDisp PTR
   DECLARE FUNCTION CreateObject (BYREF wszProgID AS WSTRING) AS IDispatch PTR
   DECLARE FUNCTION CreateObject (BYREF wszProgID AS WSTRING, BYREF wszLicKey AS WSTRING) AS IDispatch PTR
   DECLARE FUNCTION CreateObject (BYREF classID AS CONST CLSID, BYREF riid AS CONST IID) AS IDispatch PTR
   DECLARE FUNCTION CreateObject (BYREF classID AS CONST CLSID, BYREF riid AS CONST IID, BYREF wszLicKey AS WSTRING) AS IDispatch PTR
   DECLARE FUNCTION CreateObject (BYREF wszLibName AS CONST WSTRING, BYREF rclsid AS CONST CLSID, _
      BYREF riid AS CONST IID, BYREF wszLicKey AS WSTRING = "") AS IDispatch PTR
   DECLARE FUNCTION hWindow () AS HANDLE
   DECLARE SUB PixelsToHiMetric (BYVAL SizeInPix AS SIZEL PTR, BYVAL SizeInHiMetric AS SIZEL PTR)
'   DECLARE FUNCTION ForwardMessage (BYVAL pMsg AS tagMsg PTR) AS BOOLEAN
   DECLARE FUNCTION OcxDispPtr () AS IDispatch PTR
   DECLARE FUNCTION OcxDispObj () AS IDispatch PTR
   DECLARE FUNCTION Advise (BYVAL pEvtObj AS IDispatch PTR, BYVAL riid AS IID PTR) AS HRESULT
   DECLARE FUNCTION Advise (BYVAL pEvtObj AS IDispatch PTR, BYREF riid AS CONST IID) AS HRESULT
   DECLARE FUNCTION Advise (BYVAL pEvtObj AS IDispatch PTR, BYREF riid AS IID) AS HRESULT
   DECLARE FUNCTION Advise (BYVAL pEvtObj AS Afx_IDispatch PTR, BYVAL riid AS IID PTR) AS HRESULT
   DECLARE FUNCTION Advise (BYVAL pEvtObj AS Afx_IDispatch PTR, BYREF riid AS CONST IID) AS HRESULT
   DECLARE FUNCTION Advise (BYVAL pEvtObj AS Afx_IDispatch PTR, BYREF riid AS IID) AS HRESULT
   DECLARE FUNCTION Unadvise () AS HRESULT
   DECLARE STATIC FUNCTION COleConProc (BYVAL hWnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT
END TYPE
' ========================================================================================

' ========================================================================================
' Registers the class
' ========================================================================================
PRIVATE FUNCTION COleCon.RegisterClass () AS ATOM
   DIM wAtom AS ATOM
   DIM wcexw AS WNDCLASSEXW
   DIM wszClassName AS WSTRING * 260 = OC_CLASSNAME
   IF .GetClassInfoExW(.GetModuleHandleW(NULL), @wszClassName, @wcexw) = 0 THEN
      ' // Fill the WNDCLASSEXW structure
      WITH wcexw
         .cbSize        = SIZEOF(wcexw)
         .style         = CS_DBLCLKS OR CS_HREDRAW OR CS_VREDRAW
         .lpfnWndProc   = @COleConProc
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
PRIVATE SUB COleCon.CreateContainer (BYVAL pWindow AS CWindow PTR, BYVAL cID AS INTEGER, BYREF wszProgID AS WSTRING, _
   BYREF wszLibName AS WSTRING, BYREF rclsid AS const CLSID, BYREF riid AS const IID, BYREF wszLicKey AS WSTRING, _
   BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
   BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, BYVAL ambientFlags AS DWORD = 0)

   OC_DP("CONSTRUCTOR COleCon")
   ' // Initialize the COM library
   DIM hr AS HRESULT = OleInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitOLE = TRUE
   OC_DP("CONSTRUCTOR COleCon - OleInitialize hr = " & WSTR(hr))
   ' // Register the class
   DIM wszClassName AS WSTRING * 260 = OC_CLASSNAME
   this.RegisterClass
   ' // Create an instance of the OLE container Dispatch object
   m_pData = NEW Afx_OleConDispatch
   IF m_pData THEN
      ' // Store values
      m_pData->m_AmbientDlControl = ambientFlags
      m_pData->m_rxRatio = pWindow->rxRatio
      m_pData->m_ryRatio = pWindow->ryRatio
      m_pData->m_wszProgID = wszProgID     ' // ActiveX ProgID
      m_pData->m_wszLibName = wszLibName   ' // ActiveX path
      m_pData->m_wszLicKey = wszLicKey     ' // Active license key
      m_pData->m_rclsid = rclsid           ' // ActiveX CLSID
      m_pData->m_riid = riid               ' // ActiveX IID
      IF LEN(wszLibName) THEN m_pData->m_bIsRegFree = TRUE
      IF LEN(wszLicKey) THEN m_pData->m_bIsLicensed = TRUE
   END IF

   ' // Create the control
   IF dwStyle = 0 THEN dwStyle = WS_VISIBLE OR WS_TABSTOP
   IF pWindow THEN m_pData->m_hwnd = pWindow->AddControl(wszClassName, pWindow->hWindow, cID, "", _
      x, y, nWidth, nHeight, dwStyle, dwExStyle, CAST(LONG_PTR, @this))
   IF m_pData->m_hwnd THEN
      .SetWindowLongPtrW m_pData->m_hwnd, 0, CAST(LONG_PTR, @this)
      ' // Set the same font used by the parent
      DIM lfw AS LOGFONTW
      IF pWindow->Font THEN
         IF .GetObjectW(pWindow->Font, SIZEOF(lfw), @lfw) THEN m_pData->m_hFont = CreateFontIndirectW(@lfw)
      END IF
   END IF

END SUB
' ========================================================================================

' ========================================================================================
' COleCon class constructors
' ========================================================================================
PRIVATE CONSTRUCTOR COleCon (BYVAL pWindow AS CWindow PTR, BYVAL cID AS INTEGER, BYREF wszLibName AS WSTRING, _
BYREF rclsid AS const CLSID, BYREF riid AS const IID, BYREF wszLicKey AS WSTRING, _
BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, BYVAL ambientFlags AS DWORD = 0)
   this.CreateContainer(pWindow, cID, "", wszLibName, rclsid, riid, wszLicKey, x, y, nWidth, nHeight, dwStyle, dwExStyle, ambientFlags)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR COleCon (BYVAL pWindow AS CWindow PTR, BYVAL cID AS INTEGER, BYREF wszLibName AS WSTRING, _
BYREF rclsid AS const CLSID, BYREF riid AS const IID, BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, _
BYVAL nHeight AS LONG = 0, BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, BYVAL ambientFlags AS DWORD = 0)
   this.CreateContainer(pWindow, cID, "", wszLibName, rclsid, riid, "", x, y, nWidth, nHeight, dwStyle, dwExStyle, ambientFlags)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR COleCon (BYVAL pWindow AS CWindow PTR, BYVAL cID AS INTEGER, BYREF wszProgID AS WSTRING, _
BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, BYVAL ambientFlags AS DWORD = 0)
   this.CreateContainer(pWindow, cID, wszProgID, "", GUID_NULL, GUID_NULL, "", x, y, nWidth, nHeight, dwStyle, dwExStyle, ambientFlags)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR COleCon (BYVAL pWindow AS CWindow PTR, BYVAL cID AS INTEGER, BYREF wszProgID AS WSTRING, _
BYREF wszLicKey AS WSTRING, BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, BYVAL ambientFlags AS DWORD = 0)
   this.CreateContainer(pWindow, cID, wszProgID, "", GUID_NULL, GUID_NULL, wszLicKey, x, y, nWidth, nHeight, dwStyle, dwExStyle, ambientFlags)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR COleCon (BYVAL pWindow AS CWindow PTR, BYVAL cID AS INTEGER, BYREF classID AS CONST CLSID, _
BYREF riid AS CONST IID, BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, _
BYVAL nHeight AS LONG = 0, BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, BYVAL ambientFlags AS DWORD = 0)
   this.CreateContainer(pWindow, cID, "", "", classID, riid, "", x, y, nWidth, nHeight, dwStyle, dwExStyle, ambientFlags)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR COleCon (BYVAL pWindow AS CWindow PTR, BYVAL cID AS INTEGER, BYREF classID AS CONST CLSID, _
BYREF riid AS CONST IID, BYREF wszLicKey AS WSTRING, BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, _
BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, BYVAL dwStyle AS DWORD = 0, _
BYVAL dwExStyle AS DWORD = 0, BYVAL ambientFlags AS DWORD = 0)
   this.CreateContainer(pWindow, cID, "", "", classID, riid, wszLicKey, x, y, nWidth, nHeight, dwStyle, dwExStyle, ambientFlags)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' COleCon class destructor
' ========================================================================================
PRIVATE DESTRUCTOR COleCon
   OC_DP("DESTRUCTOR COleCon")
   ' // Delete the Dispatch object
   ' // (it will destroy and/or release the other objects and interfaces)
   IF m_pData THEN Delete m_pData
   ' // Free the COM library
   OC_DP("DESTRUCTOR COleCon - m_bUninitCOM = " & WSTR(m_bUninitOLE))
   IF m_bUninitOLE THEN OleUninitialize
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the handle of the container's window
' ========================================================================================
PRIVATE FUNCTION COleCon.hWindow () AS HANDLE
   FUNCTION = m_pData->m_hwnd
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts from Pixel to HiMetric
' Note: Himetric is a scaling unit similar to twips used in computing. It is one
' thousandth of a centimeter and is independent of the screen resolution.
' HIMETRIC_PER_INCH = 2540   ' 1 inch = 2.54 mm
' ========================================================================================
PRIVATE SUB COleCon.PixelsToHiMetric (BYVAL SizeInPix AS SIZEL PTR, BYVAL SizeInHiMetric AS SIZEL PTR)

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
' ========================================================================================
PRIVATE FUNCTION COleCon.OleCreateFontDisp ( _
   BYREF wszFontName AS WSTRING, _            ' __in  Font name
   BYVAL cySize AS LONGLONG, _                ' __in  Point size
   BYVAL fWeight AS SHORT, _                  ' __in  Weight (%FW_NORMAL, etc.)
   BYVAL fCharset AS SHORT, _                 ' __in  Character set
   BYVAL fItalic AS WINBOOL, _                ' __in  Italic state
   BYVAL fUnderline AS WINBOOL, _             ' __in  Underline state
   BYVAL fStrikethrough AS WINBOOL _          ' __in  Strikethrough state
   ) AS IFontDisp PTR

   OC_DP("OleCreateFontDisp " & wszFontName)
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
' Creates an instance of an object.
' Parameters:
' - wszProgID: Required. The ProgID or the CLSID of the object to create.
'   - A ProgID such as "MSCAL.Calendar.7"
'   - A CLSID such as "{8E27C92B-1264-101C-8A2F-040224009C02}"
' Return value: The dispatch interface of the control.
' ========================================================================================
PRIVATE FUNCTION COleCon.CreateObject (BYREF wszProgID AS WSTRING) AS IDispatch PTR

   OC_DP("COleCon.CreateObject - ProgID: " & wszProgID)
   ' // Exit if wszProgID is a null string
   IF LEN(wszProgID) = 0 THEN RETURN NULL
   ' // Convert the ProgID in a CLSID
   DIM ClassID AS CLSID
   CLSIDFromProgID(@wszProgID, @ClassID)
   ' // If it fails, see if it is a CLSID
   IF IsEqualIID(@ClassID, @IID_NULL) THEN
      ClassID = this.m_pData->GuidFromStr(@wszProgID)
   END IF
'   ' // If not a valid ProgID or CLSID return NULL
   IF IsEqualIID(@ClassID, @IID_NULL) THEN RETURN NULL

   ' // Create an instance of the object
   DIM pDisp AS IDispatch PTR
   DIM hr AS HRESULT = CoCreateInstance(@ClassID, NULL, CLSCTX_SERVER, @IID_IDispatch, cast(LPVOID, @pDisp))
   OC_DP("COleCon.CreateObject - hr - " & STR(hr) & " - IDispatch = " & WSTR(pDisp))
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
PRIVATE FUNCTION COleCon.CreateObject (BYREF wszProgID AS WSTRING, BYREF wszLicKey AS WSTRING) AS IDispatch PTR

   OC_DP("COleCon.CreateObject - ProgID: " & wszProgID)
   ' // Exit if wszProgID is a null string
   IF LEN(wszProgID) = 0 THEN RETURN NULL
   ' // Convert the ProgID in a CLSID
   DIM ClassID  AS CLSID
   CLSIDFromProgID(@wszProgID, @ClassID)
   ' // If it fails, see if it is a CLSID
   IF IsEqualIID(@ClassID, @IID_NULL) THEN ClassID = this.m_pData->GuidFromStr(@wszProgID)
'   ' // If not a valid CLSID return NULL
   IF IsEqualIID(@ClassID, @IID_NULL) THEN RETURN NULL

   ' // Get a reference to the IClassFactory2 interface of the control
   DIM pIClassFactory2 AS IClassFactory2 PTR, pDisp AS IDispatch PTR
   DIM hr AS HRESULT = CoGetClassObject(@ClassID, CLSCTX_SERVER, NULL, @IID_IClassFactory2, @pIClassFactory2)
   OC_DP("CoGetClassObject - hr - " & STR(hr) & " - IClassFactory2 = " & WSTR(pIClassFactory2))
   IF hr = S_OK AND pIClassFactory2 <> NULL THEN
      ' // Create a licensed instance of the control
      hr = pIClassFactory2->lpvtbl->CreateInstanceLic(pIClassFactory2, NULL, NULL, @IID_IDispatch, @wszLicKey, @pDisp)
      OC_DP("CreateInstanceLic - hr - " & STR(hr) & " - IDispatch = " & WSTR(pDisp))
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
PRIVATE FUNCTION COleCon.CreateObject (BYREF classID AS CONST CLSID, BYREF riid AS CONST IID) AS IDispatch PTR
   OC_DP("COleCon.CreateObject - classID & riid")
   ' // Create an instance of the object
   DIM pDisp AS IDIspatch PTR
   DIM hr AS HRESULT = CoCreateInstance(@ClassID, NULL, CLSCTX_SERVER, @riid, cast(LPVOID, @pDisp))
   OC_DP("COleCon.CreateObject - hr - " & STR(hr) & " - IDIspatch = " & WSTR(pDisp))
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
PRIVATE FUNCTION COleCon.CreateObject (BYREF classID AS CONST CLSID, BYREF riid AS CONST IID, BYREF wszLicKey AS WSTRING) AS IDispatch PTR
   OC_DP("COleCon.CreateObject - classID & riid")
   ' // Get a reference to the IClassFactory2 interface of the control
   DIM pIClassFactory2 AS IClassFactory2 PTR, pDisp AS IDispatch PTR
   DIM hr AS HRESULT = CoGetClassObject(@classID, CLSCTX_SERVER, NULL, @IID_IClassFactory2, @pIClassFactory2)
   OC_DP("CoGetClassObject - hr - " & STR(hr) & " - IClassFactory2 = " & WSTR(pIClassFactory2))
   IF hr = S_OK AND pIClassFactory2 <> NULL THEN
      ' // Create a licensed instance of the control
      hr = pIClassFactory2->lpvtbl->CreateInstanceLic(pIClassFactory2, NULL, NULL, @riid, @wszLicKey, @pDisp)
      OC_DP("CreateInstanceLic - hr - " & STR(hr) & " - IDispatch = " & WSTR(pDisp))
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
PRIVATE FUNCTION COleCon.CreateObject (BYREF wszLibName AS CONST WSTRING, BYREF rclsid AS CONST CLSID, _
   BYREF riid AS CONST IID, BYREF wszLicKey AS WSTRING = "") AS IDispatch PTR

   ' // See if the library is already loaded in the address space
   DIM hLib AS HANDLE = GetModuleHandleW(wszLibName)
   ' // If it is not loaded, load it
   IF hLib = NULL THEN hLib = DyLibLoad(wszLibName)
   ' // If it fails, abort
   IF hLib = NULL THEN RETURN NULL

   ' // Retrieve the address of the exported function DllGetClassObject
   DIM pfnDllGetClassObject AS FUNCTION (BYVAL rclsid AS CONST IID CONST PTR, BYVAL riid AS CONST IID CONST PTR, BYVAL ppv AS LPVOID PTR) AS HRESULT
   pfnDllGetClassObject = DyLibSymbol(hLib, "DllGetClassObject")
   IF pfnDllGetClassObject = NULL THEN RETURN NULL

   IF LEN(wszLicKey) = 0 THEN
      ' // Request a reference to the IClassFactory interface
      DIM pIClassFactory AS IClassFactory PTR, pDisp AS IDispatch PTR
      DIM hr AS HRESULT = pfnDllGetClassObject(@rclsid, @IID_IClassFactory, @pIClassFactory)
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
      IF hr <> S_OK THEN RETURN NULL
      ' // Create a licensed instance of the server or control
      hr = pIClassFactory2->lpVtbl->CreateInstanceLic(pIClassFactory2, NULL, NULL, @riid, @wszLicKey, @pDisp)
      pIClassFactory2->lpVtbl->Release(pIClassFactory2)
      IF hr <> S_OK THEN RETURN NULL
      RETURN pDisp
   END IF

END FUNCTION
' ========================================================================================

' ========================================================================================
' OLE container control window procedure
' ========================================================================================
PRIVATE FUNCTION COleCon.COleConProc (BYVAL hWnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   DIM hr AS HRESULT

   SELECT CASE uMsg

      CASE WM_CREATE
         ' // Get a pointer to the OLE Container class
         DIM pCreateStruct AS CREATESTRUCT PTR = CAST(CREATESTRUCT PTR, lParam)
         DIM pOleCon AS COleCon PTR = CAST(COleCon PTR, pCreateStruct->lpCreateParams)
         IF pOleCon = NULL THEN RETURN -1
         ' // Store data in Afx_OleConDispatch
         pOleCon->m_pData->m_hwnd = hwnd
         ' // Create an instance of the ActiveX control
         IF LEN(pOleCon->m_pData->m_wszLibName) THEN
            pOleCon->m_pData->m_pOcx = pOleCon->CreateObject(pOleCon->m_pData->m_wszLibName, _
               pOleCon->m_pData->m_rclsid, pOleCon->m_pData->m_riid, pOlecon->m_pData->m_wszLicKey)
         ELSEIF LEN(pOleCon->m_pData->m_wszProgID) <> 0 AND LEN(pOleCon->m_pData->m_wszLicKey) <> 0 THEN
            pOleCon->m_pData->m_pOcx = pOleCon->CreateObject(pOleCon->m_pData->m_wszProgID, pOleCon->m_pData->m_wszLicKey)
         ELSEIF LEN(pOleCon->m_pData->m_wszProgID) <> 0 THEN
            pOleCon->m_pData->m_pOcx = pOleCon->CreateObject(pOleCon->m_pData->m_wszProgID)
         ELSEIF IsEqualIID(@pOleCon->m_pData->m_rclsid, @IID_NULL) = FALSE AND IsEqualIID(@pOleCon->m_pData->m_riid, @IID_NULL) = FALSE AND LEN(pOlecon->m_pData->m_wszLicKey) <> 0 THEN
            pOleCon->m_pData->m_pOcx = pOleCon->CreateObject(pOleCon->m_pData->m_rclsid, pOleCon->m_pData->m_riid, pOlecon->m_pData->m_wszLicKey)
         ELSEIF IsEqualIID(@pOleCon->m_pData->m_rclsid, @IID_NULL) = FALSE AND IsEqualIID(@pOleCon->m_pData->m_riid, @IID_NULL) = FALSE THEN
            pOleCon->m_pData->m_pOcx = pOleCon->CreateObject(pOleCon->m_pData->m_rclsid, pOleCon->m_pData->m_riid)
         END IF
        ' // If it fails, abort the creation of the control
         IF pOleCon->m_pData->m_pOcx = NULL THEN RETURN -1
         ' // Creates a default font
         IF AfxWindowsVersion >= 600 AND CLNG(AfxIsProcessDpiAware) THEN
            pOleCon->m_pData->m_pIFontDisp = pOleCon->OleCreateFontDisp("Segoe UI", 9, FW_NORMAL, ANSI_CHARSET, 0, 0, 0)
         ELSE
            pOleCon->m_pData->m_pIFontDisp = pOleCon->OleCreateFontDisp("Tahoma", 8, FW_NORMAL, ANSI_CHARSET, 0, 0, 0)
         END IF
         ' // Retrieves the IOleObject interface
         IUnknown_QueryInterface(pOleCon->m_pData->m_pOcx, @IID_IOleObject, @pOleCon->m_pData->m_pIOleObject)
         ' // If it doesn't exist, it's not a visual control
         IF pOleCon->m_pData->m_pIOleObject = NULL THEN RETURN -1
         ' // Gets the status bitwise flags
         pOleCon->m_pData->m_pIOleObject->GetMiscStatus(DVASPECT_CONTENT, @pOlecon->m_pData->m_dwMiscStatus)
         ' // The control needs to be put on running state
         IF (pOlecon->m_pData->m_dwMiscStatus AND OLEMISC_ALWAYSRUN) = OLEMISC_ALWAYSRUN THEN
            hr = OleRun(cast(IUnknown PTR, pOlecon->m_pData->m_pOcx))
            OC_DP("OleRun: " & HEX(hr))
         END IF
         IF (pOlecon->m_pData->m_dwMiscStatus AND OLEMISC_SIMPLEFRAME) = OLEMISC_SIMPLEFRAME THEN
            OC_DP("--- Simple frame ---")
         END IF
         ' // The control wants that you set the client site first
         IF (pOlecon->m_pData->m_dwMiscStatus AND OLEMISC_SETCLIENTSITEFIRST) = OLEMISC_SETCLIENTSITEFIRST THEN
            ' // Sets the client site
            OC_DP("--- SetClientSite - Begin ---")
            IF pOlecon->m_pData->m_pIOleClientSiteImpl THEN
               DIM pIOleClientSite AS ANY PTR = pOlecon->m_pData->m_pIOleClientSiteImpl
               hr = pOlecon->m_pData->m_pIOleObject->SetClientSite(pIOleClientSite)
               OC_DP("--- SetClientSite - End --- " & HEX(hr))
            END IF
         END IF
         ' // Sets the host name
         DIM wszContainerApp AS WSTRING * 260 = OC_CLASSNAME
         hr = pOlecon->m_pData->m_pIOleObject->SetHostNames(@wszContainerApp, NULL)
         OC_DP("SetHostNames: " & HEX(hr))
         ' // Initializes the object to a default state
         DIM psi AS IPersistStreamInit PTR
         hr = IUnknown_QueryInterface(pOleCon->m_pData->m_pOcx, @IID_IPersistStreamInit, @psi)
         IF hr = S_OK AND psi <> NULL THEN
            hr = psi->lpvtbl->InitNew(psi)
            OC_DP("IPersistStreamInit.InitNew - hr: " & HEX(hr))
            IUnknown_Release(psi)
         END IF
         IF (pOlecon->m_pData->m_dwMiscStatus AND OLEMISC_SETCLIENTSITEFIRST) = OLEMISC_SETCLIENTSITEFIRST THEN
            ' // Sets the client site
            OC_DP("--- SetClientSite - Begin ---")
            IF pOlecon->m_pData->m_pIOleClientSiteImpl THEN
               DIM pIOleClientSite AS ANY PTR = pOlecon->m_pData->m_pIOleClientSiteImpl
               hr = pOlecon->m_pData->m_pIOleObject->SetClientSite(pIOleClientSite)
               OC_DP("--- SetClientSite - End --- " & HEX(hr))
            END IF
         END IF
         ' // Gets the size of the container's window
         DIM rc AS RECT
         GetClientRect pOleCon->m_pData->m_hwnd, @rc
         ' // Converts from pixel to himetric
         DIM pxSize AS SIZEL, hmSize AS SIZEL
         pxSize.cx = rc.Right - rc.Left
         pxSize.cy = rc.Bottom - rc.Top
         pOleCon->PixelsToHiMetric(@pxSize, @hmSize)
         ' // Sets extent of object's display area
         OC_DP("--- DoVerb - INPLACEACTIVATE - Begin ---")
         IF pOlecon->m_pData->m_pIOleObject THEN pOleCon->m_pData->m_pIOleObject->SetExtent(DVASPECT_CONTENT, @hmSize)
         ' // Activates the control in place
         IF pOleCon->m_pData->m_pIOleClientSiteImpl THEN
            DIM pIOleClientSite AS ANY PTR = pOleCon->m_pData->m_pIOleClientSiteImpl
            IF pOlecon->m_pData->m_pIOleObject THEN
               DIM hr AS HRESULT = pOlecon->m_pData->m_pIOleObject->DoVerb(OLEIVERB_INPLACEACTIVATE, NULL, _
                  cast(IOleClientSite PTR, pIOleClientSite), 0, pOleCon->m_pData->m_hwnd, @rc)
            END IF
            OC_DP("--- DoVerb - INPLACEACTIVATE - End --- " & HEX(hr))
            ' // Sets extent of object's display area
            IF pOlecon->m_pData->m_pIOleObject THEN pOleCon->m_pData->m_pIOleObject->SetExtent(DVASPECT_CONTENT, @hmSize)
            ' // Indicates how much of the in-place object is visible.
            IF pOleCon->m_pData->m_pIOleInPlaceObject = NULL THEN
               IUnknown_QueryInterface(pOleCon->m_pData->m_pOcx, @IID_IOleInPlaceObject, @pOlecon->m_pData->m_pIOleInPlaceObject)
            END IF
            IF pOleCon->m_pData->m_pIOleInPlaceObject THEN
               pOleCon->m_pData->m_pIOleInPlaceObject->SetObjectRects(@rc, @rc)
            END IF
         END IF
         EXIT FUNCTION

      CASE WM_ERASEBKGND
         ' // Don't erase the background to avoid flicker
         FUNCTION = 1
         EXIT FUNCTION

   	CASE WM_ENABLE
         InvalidateRect hwnd, NULL, 0
         UpdateWindow hwnd
         EXIT FUNCTION

      CASE WM_SETFOCUS
         DIM pOleCon AS COleCon PTR = CAST(COleCon PTR, GetWindowLongPtrW(hwnd, 0))
         IF pOleCon = NULL THEN EXIT FUNCTION
         pOleCon->m_pData->m_bHaveFocus = TRUE
         ' // Activates an object in place, along with its full set of user-interface tools.
         IF pOleCon->m_pData->m_pIOleObject <> NULL AND pOleCon->m_pData->m_bInPlaceActive = FALSE THEN
            DIM rc AS RECT
            GetClientRect hwnd, @rc
            IF pOleCon->m_pData->m_pIOleClientSiteImpl THEN
               DIM pIOleClientSite AS ANY PTR = pOleCon->m_pData->m_pIOleClientSiteImpl
               pOleCon->m_pData->m_pIOleObject->DoVerb(OLEIVERB_UIACTIVATE, NULL, pIOleClientSite, 0, hwnd, @rc)
            END IF
         END IF
         IF pOleCon->m_pData->m_pIOleInPlaceObject = NULL THEN
            IUnknown_QueryInterface(pOleCon->m_pData->m_pOcx, @IID_IOleInPlaceObject, @pOleCon->m_pData->m_pIOleInPlaceObject)
         END IF
         IF pOleCon->m_pData->m_pIOleInPlaceObject THEN
            DIM hwndCtrl AS HANDLE
            pOleCon->m_pData->m_pIOleInPlaceObject->GetWindow(@hwndCtrl)
            IF IsWindow(hwndCtrl) THEN SetFocus hwndCtrl
         END IF
         EXIT FUNCTION

      CASE WM_KILLFOCUS
         OC_DP("--- WM_KILLFOCUS BEGIN ---")
         DIM pOleCon AS COleCon PTR = CAST(COleCon PTR, GetWindowLongPtrW(hwnd, 0))
         IF pOleCon = NULL THEN EXIT FUNCTION
         pOleCon->m_pData->m_bHaveFocus = FALSE
         OC_DP("--- WM_KILLFOCUS END ---")
         EXIT FUNCTION

      ' ==================================================================================
      ' Resizes the control to fill the client area of the host window
      ' ==================================================================================
      CASE WM_SIZE
         OC_DP("--- WM_SIZE BEGIN ---")
         DIM pOleCon AS COleCon PTR = CAST(COleCon PTR, GetWindowLongPtrW(hwnd, 0))
         IF pOleCon = NULL THEN EXIT FUNCTION
         IF pOleCon->m_pData->m_bInPlaceActive THEN
            ' // Retrieves the coordinates of a window's client area.
            DIM rc AS RECT
            GetClientRect hwnd, @rc
            ' // Converts from pixel to himetric
            DIM pxSize AS SIZEL, hmSize AS SIZEL
            pxSize.cx = rc.Right - rc.Left
            pxSize.cy = rc.Bottom - rc.Top
            pOleCon->PixelsToHiMetric(@pxSize, @hmSize)
            ' // Sets extent of object's display area
            IF pOlecon->m_pData->m_pIOleObject THEN pOleCon->m_pData->m_pIOleObject->SetExtent(DVASPECT_CONTENT, @hmSize)
            ' // Indicates how much of the in-place object is visible.
            IF pOleCon->m_pData->m_pIOleInPlaceObject = NULL THEN
               IUnknown_QueryInterface(pOleCon->m_pData->m_pOcx, @IID_IOleInPlaceObject, @pOlecon->m_pData->m_pIOleInPlaceObject)
            END IF
            IF pOleCon->m_pData->m_pIOleInPlaceObject THEN
               pOleCon->m_pData->m_pIOleInPlaceObject->SetObjectRects(@rc, @rc)
            END IF
         END IF
         OC_DP("--- WM_SIZE_END ---")
         EXIT FUNCTION

   END SELECT

   ' // Default processing for other messages.
   FUNCTION = DefWindowProcW(hWnd, uMsg, wParam, lParam)

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
' Note: Use the AfxForwardMessage function below that allows a more generic usage.
' ========================================================================================
'PRIVATE FUNCTION COleCon.ForwardMessage (BYVAL pMsg AS tagMsg PTR) AS BOOLEAN
'   IF m_pData->m_pOcx = NULL THEN RETURN FALSE
'   ' // See if the window with focus is the WebBrowser control
'   IF m_pData->m_wszProgID = "Shell.Explorer" OR m_pData->GuidText(@m_pData->m_rclsid) = "8856F961-340A-11D0-A96B-00C04FD705A2"  THEN
'      DIM hwndFocus AS HWND = GetFocus
'      DIM hWin AS HWND = FindWindowEx(m_pData->m_hwnd, NULL, "Shell Embedding", "")
'      IF hWin THEN hWin = FindWindowEx(hWin, NULL, "Shell DocObject View", "")
'      IF hWin THEN hWin = FindWindowEx(hWin, NULL, "Internet Explorer_Server", "")
'      IF hWin <> hwndFocus THEN RETURN FALSE
'   ELSE
'      ' // See if the window that has the focus is a child of our container window
'      DIM hWndParent AS HWND, hWndTmp AS HWND
'      hWndParent = GetFocus()
'      DO
'         hWndTmp = GetParent(hWndParent)
'         IF hWndTmp = NULL THEN EXIT DO
'         hWndParent = hWndTmp
'         IF hWndParent = m_pData->m_hwnd THEN EXIT DO
'      LOOP
'      IF hWndParent <> m_pData->m_hwnd THEN RETURN FALSE
'   END IF
'   ' // Translate the message
'   DIM hr AS HRESULT, pActiveObject AS IOleInPlaceActiveObject PTR
'   hr = IUnknown_QueryInterface(m_pData->m_pOcx, @IID_IOleInPlaceActiveObject, @pActiveObject)
'   IF pActiveObject = NULL THEN RETURN FALSE
'   hr = pActiveObject->lpvtbl->TranslateAccelerator(pActiveObject, pMsg)
'   IUnknown_Release(pActiveObject)
'   IF hr = S_OK THEN RETURN TRUE
'END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a reference to the hosted ActiveX control dispatch interface given the handle of
' the window that hosts it. Don't call IUnknown_Release.
' ========================================================================================
PRIVATE FUNCTION COleCon.OcxDispPtr () AS IDispatch PTR
   RETURN m_pData->m_pOcx
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a counted reference to the hosted ActiveX control dispatch interface given the
' handle of the window that hosts it. You must call IUnknown_Release when no longer need it.
' ========================================================================================
PRIVATE FUNCTION COleCon.OcxDispObj () AS IDispatch PTR
   IF m_pData->m_pOcx <> NULL THEN
      IUnknown_AddRef(m_pData->m_pOcx)
      RETURN m_pData->m_pOcx
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Establishes a connection between the connection point object and the client's sink.
' ========================================================================================
PRIVATE FUNCTION COleCon.Advise (BYVAL pEvtObj AS IDispatch PTR, BYVAL riid AS IID PTR) AS HRESULT
   OC_DP("BEGIN COlecon.Advise")
   IF pEvtObj = NULL THEN RETURN E_POINTER
   m_riidEvt = riid
   ' // Query for the IConnectionPointContainer interface
   DIM pCPC AS IConnectionPointContainer PTR
   DIM hr AS HRESULT = IUnknown_QueryInterface(m_pData->m_pOcx, @IID_IConnectionPointContainer, @pCPC)
   IF hr <> S_OK OR pCPC = NULL THEN RETURN hr
   ' // Query for the IConnectionPoint interface
   DIM pCP AS IConnectionPoint PTR
   hr = pCPC->lpvtbl->FindConnectionPoint(pCPC, riid, @pCP)
   IF hr <> S_OK OR pCP = NULL THEN
      IUnknown_Release(pCPC)
      RETURN hr
   END IF
   ' // Terminates the advisory connection previously established between a connection point object and a client's sink.
   IF m_dwCookie THEN hr = pCP->lpvtbl->Unadvise(pCP, m_dwCookie)
   m_dwCookie = 0
   ' // Establishes a connection between a connection point object and the client's sink.
   hr = pCP->lpvtbl->Advise(pCP, cast(IUnknown PTR, pEvtObj), @m_dwCookie)
   IF pCPC THEN IUnknown_Release(pCPC)
   IF pCP THEN IUnknown_Release(pCP)
   OC_DP("END COlecon.Advise")
   RETURN hr
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION COleCon.Advise (BYVAL pEvtObj AS IDispatch PTR, BYREF riid AS CONST IID) AS HRESULT
   RETURN this.Advise(pEvtObj, cast(IID PTR, @riid))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION COleCon.Advise (BYVAL pEvtObj AS IDispatch PTR, BYREF riid AS IID) AS HRESULT
   RETURN this.Advise(pEvtObj, cast(IID PTR, @riid))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION COleCon.Advise (BYVAL pEvtObj AS Afx_IDispatch PTR, BYVAL riid AS IID PTR) AS HRESULT
   RETURN this.Advise(cast(IDispatch PTR, cast(ULONG_PTR, pEvtObj)), riid)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION COleCon.Advise (BYVAL pEvtObj AS Afx_IDispatch PTR, BYREF riid AS CONST IID) AS HRESULT
   RETURN this.Advise(cast(IDispatch PTR, cast(ULONG_PTR, pEvtObj)), cast(IID PTR, @riid))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION COleCon.Advise (BYVAL pEvtObj AS Afx_IDispatch PTR, BYREF riid AS IID) AS HRESULT
   RETURN this.Advise(cast(IDispatch PTR, cast(ULONG_PTR, pEvtObj)), cast(IID PTR, @riid))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Releases the events connection identified with the cookie returned by the Advise method.
' ========================================================================================
PRIVATE FUNCTION COleCon.Unadvise () AS HRESULT
   OC_DP("BEGIN COlecon.Unadvise")
   ' // Not a valid connection
   IF m_dwCookie = 0 THEN RETURN E_POINTER
   ' // Query for the IConnectionPointContainer interface
   DIM pCPC AS IConnectionPointContainer PTR
   DIM hr AS HRESULT = m_pData->m_pOcx->lpvtbl->QueryInterface(m_pData->m_pOcx, @IID_IConnectionPointContainer, @pCPC)
   IF hr <> S_OK OR pCPC = NULL THEN RETURN hr
   ' // Query for the IConnectionPoint interface
   DIM pCP AS IConnectionPoint PTR
   hr = pCPC->lpvtbl->FindConnectionPoint(pCPC, m_riidEvt, @pCP)
   IF hr <> S_OK OR pCP = NULL THEN
      IUnknown_Release(pCPC)
      RETURN hr
   END IF
   ' // Terminates the advisory connection previously established between a connection point object and a client's sink.
   hr = pCP->lpvtbl->Unadvise(pCP, m_dwCookie)
   m_dwCookie = 0
   IF pCPC THEN IUnknown_Release(pCPC)
   IF pCP THEN IUnknown_Release(pCP)
   OC_DP("END COlecon.Unadvise")
   RETURN hr
END FUNCTION
' ========================================================================================

' ########################################################################################
'                              *** HELPER PROCEDURES ***
' ########################################################################################

' ========================================================================================
' Returns the OLE container window handle given the handle of the form or any control in the
' form and its control identifier.
' Parameters:
' - hwnd = Reference window handle.
' - Control identifier, e.g. IDC_WEBBROWSER.
' Returns the OLE container window handle or NULL.
' ========================================================================================
PRIVATE FUNCTION AfxCOleConWindow (BYVAL hwnd AS HWND, BYVAL cID AS WORD) AS HWND
   DIM wszClassName AS WSTRING * 260
   DIM hwndChild AS .HWND = AfxGetControlHandle(hwnd, cID)
   GetClassNameW hwndChild, wszClassName, SIZEOF(wszClassName)
   IF wszClassName <> OC_CLASSNAME THEN RETURN NULL
   RETURN hwndChild
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a pointer to the COleCon class given the handle of its associated window.
' ========================================================================================
PRIVATE FUNCTION AfxCOleConPtr OVERLOAD (BYVAL hwnd AS HWND) AS COleCon PTR
   DIM wszClassName AS WSTRING * 260
   GetClassNameW hwnd, wszClassName, SIZEOF(wszClassName)
   IF wszClassName <> OC_CLASSNAME THEN RETURN NULL
   RETURN CAST(COleCon PTR, .GetWindowLongPtrW(hwnd, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a pointer to the COleCon class given the handle of the form or any control in the
' form and the control identifier.
' Parameters:
' - hwnd = Reference window handle.
' - Control identifier, e.g. IDC_WEBBROWSER.
' Returns a pointer to the COleCon class or NULL.
' ========================================================================================
PRIVATE FUNCTION AfxCOleConPtr OVERLOAD (BYVAL hwnd AS HWND, BYVAL cID AS WORD) AS COleCon PTR
   DIM wszClassName AS WSTRING * 260
   DIM hwndChild AS .HWND = AfxGetControlHandle(hwnd, cID)
   GetClassNameW hwndChild, wszClassName, SIZEOF(wszClassName)
   IF wszClassName <> OC_CLASSNAME THEN RETURN NULL
   RETURN CAST(COleCon PTR, .GetWindowLongPtrW(hwndChild, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a reference to the hosted ActiveX control dispatch interface given the handle of
' the window that hosts it. Don't call IUnknown_Release.
' ========================================================================================
PRIVATE FUNCTION AfxCOleConDispPtr (BYVAL hwnd AS HWND) AS IDispatch PTR
   DIM wszClassName AS WSTRING * 260
   GetClassNameW hwnd, wszClassName, SIZEOF(wszClassName)
   IF wszClassName <> OC_CLASSNAME THEN RETURN NULL
   DIM pOleCon AS COleCon PTR = CAST(COleCon PTR, .GetWindowLongPtrW(hwnd, 0))
   IF pOleCon THEN RETURN pOleCon->m_pData->m_pOcx
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns a counted reference to the hosted ActiveX control dispatch interface given the
' handle of the window that hosts it. You must call IUnknown_Release when no longer need it.
' ========================================================================================
PRIVATE FUNCTION AfxCOleConDispObj (BYVAL hwnd AS HWND) AS IDispatch PTR
   DIM wszClassName AS WSTRING * 260
   GetClassNameW hwnd, wszClassName, SIZEOF(wszClassName)
   IF wszClassName <> OC_CLASSNAME THEN RETURN NULL
   DIM pOleCon AS COleCon PTR = CAST(COleCon PTR, .GetWindowLongPtrW(hwnd, 0))
   IF pOleCon THEN
      IF pOleCon->m_pData->m_pOcx <> NULL THEN
         IUnknown_AddRef(pOleCon->m_pData->m_pOcx)
         RETURN pOleCon->m_pData->m_pOcx
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
'      IF AfxForwardMessage(GetFocus, @uMsg) = FALSE THEN
'         IF IsDialogMessageW(hWndMain, @uMsg) = 0 THEN
'            TranslateMessage(@uMsg)
'            DispatchMessageW(@uMsg)
'         END IF
'      END IF
'   WEND
'   FUNCTION = uMsg.wParam
' ========================================================================================
PRIVATE FUNCTION AfxForwardMessage (BYVAL hctl AS HWND, BYVAL pMsg AS tagMsg PTR) AS BOOLEAN
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
         IF wszClassName = OC_CLASSNAME THEN EXIT DO
      LOOP
      IF wszClassName <> OC_CLASSNAME THEN RETURN FALSE
      ' // Get a pointer to the OLE container class
      DIM pOleCon AS COleCon PTR = cast(COleCon PTR, GetWindowLongPtrW(hwndParent, 0))
      IF pOleCon = NULL THEN RETURN FALSE
      IF pOlecon->m_pData->m_pOcx = NULL THEN RETURN FALSE
      ' // Translate the message
      DIM hr AS HRESULT, pActiveObject AS IOleInPlaceActiveObject PTR
      hr = IUnknown_QueryInterface(pOleCon->m_pData->m_pOcx, @IID_IOleInPlaceActiveObject, @pActiveObject)
      IF pActiveObject = NULL THEN RETURN FALSE
      hr = pActiveObject->lpvtbl->TranslateAccelerator(pActiveObject, pMsg)
      IUnknown_Release(pActiveObject)
      IF hr = S_OK THEN RETURN TRUE
   END IF
   RETURN FALSE
END FUNCTION
' ========================================================================================
