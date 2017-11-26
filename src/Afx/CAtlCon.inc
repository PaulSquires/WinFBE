' ########################################################################################
' Microsoft Windows
' File: CAtlCon.inc
' Contents: ATL OLE Container class
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2017 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "Afx/AfxAtlLib.bi"
#include once "Afx/AfxExDisp.bi"
#include once "Afx/AfxCOM.inc"
#include once "Afx/CWindow.inc"

NAMESPACE Afx

' ========================================================================================
' Macro for debug
' To allow debugging, define _CATL_DEBUG_ 1 in your application before including this file.
' ========================================================================================
#ifndef _CATL_DEBUG_
   #define _CATL_DEBUG_ 0
#endif
#ifndef _CATL_DP_
   #define _CaTL_DP_ 1
   #MACRO CATL_DP(st)
      #if (_CATL_DEBUG_ = 1)
         OutputDebugStringW(st)
      #endif
   #ENDMACRO
#endif
' ========================================================================================

' ########################################################################################
'                                *** CAtlCon class ***
' ########################################################################################

' ========================================================================================
' CAtlCon class
' Implements an OLE container as a Windows custom control.
' ========================================================================================
TYPE CAtlCon

Private:
   m_bUninitOLE AS BOOLEAN   ' // Ole initialization flag
   m_hAtlLib AS HMODULE      ' // Handle of the ATL library
   AFX_IID_IAxWinAmbientDispatch_ AS GUID = (&hB6EA2051, &h048A, &h11D1, {&h82, &hB9, &h00, &hC0, &h4F, &hB9, &h94, &h2E})
   DECLARE FUNCTION WinInit () AS BOOLEAN

Protected:
   m_wszAtlLibBame AS WSTRING * 260 = "Atl.dll"      ' // ATL dll name
   m_wszAtlClassName AS WSTRING * 260 = "AtlAxWin"   ' // ATL class name

Public:
   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL pWindow AS CWindow PTR, BYVAL cID AS INTEGER, BYREF wszName AS WSTRING = "", _
           BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
           BYVAL dwStyle AS DWORD = WS_VISIBLE OR WS_CHILD OR WS_TABSTOP, BYVAL dwExStyle AS DWORD = 0)
   DECLARE CONSTRUCTOR (BYVAL pWindow AS CWindow PTR, BYVAL cID AS INTEGER, BYREF wszProgID AS WSTRING, _
           BYREF wszLicKey AS WSTRING, BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
           BYVAL dwStyle AS DWORD = WS_VISIBLE OR WS_CHILD OR WS_TABSTOP, BYVAL dwExStyle AS DWORD = 0)
   DECLARE CONSTRUCTOR (BYVAL pWindow AS CWindow PTR, BYVAL cID AS INTEGER, BYREF wszLibName AS CONST WSTRING, _
           BYREF rclsid AS CONST CLSID, BYREF riid AS CONST IID, BYREF wszLicKey AS WSTRING = "", _
           BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
           BYVAL dwStyle AS DWORD = WS_VISIBLE OR WS_CHILD OR WS_TABSTOP, BYVAL dwExStyle AS DWORD = 0)
   DECLARE FUNCTION AddControl (BYVAL pWindow AS CWindow PTR, BYVAL cID AS INTEGER, _
           BYREF wszName AS WSTRING = "", BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, _
           BYVAL nHeight AS LONG = 0, BYVAL dwStyle AS DWORD = WS_VISIBLE OR WS_CHILD OR WS_TABSTOP, BYVAL dwExStyle AS DWORD = 0) AS HWND
   DECLARE FUNCTION AddControl (BYVAL pWindow AS CWindow PTR, BYVAL cID AS INTEGER, _
           BYREF wszProgID AS WSTRING, BYREF wszLicKey AS WSTRING, BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, _
           BYVAL nHeight AS LONG = 0, BYVAL dwStyle AS DWORD = WS_VISIBLE OR WS_CHILD OR WS_TABSTOP, BYVAL dwExStyle AS DWORD = 0) AS HWND
   DECLARE FUNCTION AddControl (BYVAL pWindow AS CWindow PTR, BYVAL cID AS INTEGER, BYREF wszLibName AS CONST WSTRING, _
           BYREF rclsid AS CONST CLSID, BYREF riid AS CONST IID, BYREF wszLicKey AS WSTRING = "", _
           BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
           BYVAL dwStyle AS DWORD = WS_VISIBLE OR WS_CHILD OR WS_TABSTOP, BYVAL dwExStyle AS DWORD = 0) AS HWND
   DECLARE FUNCTION AttachControl (BYVAL pControl AS IUnknown PTR, BYVAL hCtl AS HWND) AS HRESULT
   DECLARE FUNCTION Init () AS BOOLEAN
   DECLARE FUNCTION AtlVersion () AS DWORD
   DECLARE FUNCTION HiMetricToPixel (BYREF SizeInHiMetric AS SIZEL) AS SIZEL
   DECLARE FUNCTION PixelToHiMetric (BYREF SizeInPix AS SIZEL) AS SIZEL
   DECLARE FUNCTION GetHost (BYVAL hwnd AS HWND) AS IUnknown PTR
   DECLARE FUNCTION GetControl (BYVAL hwnd AS HWND) AS IUnknown PTR
   DECLARE FUNCTION GetDispatch (BYVAL hwnd AS .HWND) AS IDispatch PTR
   DECLARE PROPERTY DocHostFlags (BYVAL hwnd AS .HWND, BYVAL dwDocHostFlags AS DWORD)
   DECLARE PROPERTY DocHostFlags (BYVAL hwnd AS .HWND) AS DWORD
   DECLARE PROPERTY DocHostDoubleClickFlags (BYVAL hwnd AS .HWND, BYVAL dwDocHostDoubleClickFlags AS DWORD)
   DECLARE PROPERTY DocHostDoubleClickFlags (BYVAL hwnd AS .HWND) AS DWORD
   DECLARE PROPERTY AllowContextMenu (BYVAL hwnd AS .HWND, BYVAL bAllowContextMenu AS VARIANT_BOOL)
   DECLARE PROPERTY AllowContextMenu (BYVAL hwnd AS .HWND) AS VARIANT_BOOL
   DECLARE PROPERTY AllowShowUI (BYVAL hwnd AS .HWND, BYVAL bAllowShowUI AS VARIANT_BOOL)
   DECLARE PROPERTY AllowShowUI (BYVAL hwnd AS .HWND) AS VARIANT_BOOL
   DECLARE PROPERTY AllowWindowlessActivation (BYVAL hwnd AS .HWND, BYVAL bAllowWindowless AS VARIANT_BOOL)
   DECLARE PROPERTY AllowWindowlessActivation (BYVAL hwnd AS .HWND) AS VARIANT_BOOL
   DECLARE PROPERTY BackColor (BYVAL hwnd AS .HWND, BYVAL clrBackground AS OLE_COLOR)
   DECLARE PROPERTY BackColor (BYVAL hwnd AS .HWND) AS OLE_COLOR
   DECLARE PROPERTY ForeColor (BYVAL hwnd AS .HWND, BYVAL clrBackground AS OLE_COLOR)
   DECLARE PROPERTY ForeColor (BYVAL hwnd AS .HWND) AS OLE_COLOR
   DECLARE PROPERTY LocaleID (BYVAL hwnd AS .HWND, BYVAL lcidLocaleID AS LCID)
   DECLARE PROPERTY LocaleID (BYVAL hwnd AS .HWND) AS LCID
   DECLARE PROPERTY DisplayAsDefault (BYVAL hwnd AS .HWND, BYVAL bDisplayAsDefault AS VARIANT_BOOL)
   DECLARE PROPERTY DisplayAsDefault (BYVAL hwnd AS .HWND) AS VARIANT_BOOL
   DECLARE PROPERTY MessageReflect (BYVAL hwnd AS .HWND, BYVAL bMessageReflect AS VARIANT_BOOL)
   DECLARE PROPERTY MessageReflect (BYVAL hwnd AS .HWND) AS VARIANT_BOOL
   DECLARE PROPERTY Font (BYVAL hwnd AS .HWND, BYVAL pFont AS IFontDisp PTR)
   DECLARE PROPERTY Font (BYVAL hwnd AS .HWND) AS IFontDisp PTR
   DECLARE FUNCTION Advise (BYVAL punkCP AS IUnknown PTR, BYVAL punk AS IUnknown PTR, BYVAL riid AS IID PTR, BYREF dwCookie AS DWORD) AS HRESULT
   DECLARE FUNCTION Unadvise (BYVAL punkCP AS IUnknown PTR, BYVAL riid AS IID PTR, BYVAL dwCookie AS DWORD) AS HRESULT

END TYPE
' ========================================================================================

' ========================================================================================
' Initializes ATL's control hosting code by registering the "AtlAxWin" and "AtlAxWinLic"
' window classes plus a couple of custom window messages, WM_GetHost and WM_GetControl.
' AtlAxWinInit cheks if the class has already been registered. Therefore, there is no
' danger in calling it more than once.
' ========================================================================================
PRIVATE FUNCTION CAtlCon.WinInit () AS BOOLEAN
   CATL_DP("CAtlCon WinInit")
   DIM pProc AS FUNCTION () AS BOOLEAN
   pProc = CAST(ANY PTR, GetProcAddress(m_hAtlLib, "AtlAxWinInit"))
   IF pProc = NULL THEN RETURN FALSE
   RETURN pProc()
END FUNCTION
' ========================================================================================

' ========================================================================================
' Initializes the COM library and loads and initializes the ATL libary.
' ========================================================================================
PRIVATE FUNCTION CAtlCon.Init () AS BOOLEAN
   CATL_DP("CAtlCon Init")
   ' // Initialize the COM library
   DIM hr AS HRESULT = OleInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitOLE = TRUE
   ' // Load the ATL library
   m_hAtlLib = LoadLibraryW(m_wszAtlLibBame)
   IF m_hAtlLib = NULL THEN
      MessageBoxW(GetActiveWindow, "Unable to load " & m_wszAtlLibBame, "Error", MB_APPLMODAL OR MB_ICONERROR)
      RETURN FALSE
   END IF
   ' // Initialize ATL
   this.WinInit
   RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Default constructor
' Example:
'   ' // Create an instance of the WebBrowser control
'   DIM pCon AS CAtlCon
'   ' // Create an instance of the WebBrowser control
'   DIM pObj AS ANY PTR = AfxNewCom("Shell.Explorer")
'   ' // Add an AtlAxWindow
'   DIM hWb AS HWND = pCon.AddControl(@pWindow, IDC_WEBBROWSER, "", 0, 0, pWindow.ClientWidth, pWindow.ClientHeight)
'   ' // Attach the WebBrowser control to the ATL window
'   DIM hr AS HRESULT = pCon.AttachControl(pObj, hWb)
'   ' // Make the web page themed and DPI aware
'   pCon.DocHostFlags(hWb) = DOCHOSTUIFLAG_NO3DBORDER OR DOCHOSTUIFLAG_THEME OR DOCHOSTUIFLAG_DPI_AWARE
'   ' // Get a direct pointer to the Afx_IWebBrowser2 interface
'   DIM pWb AS Afx_IWebBrowser2 PTR = AtlBrowserDispatch(hWb)
'   ' // Make navigation silent (disables annoying popups of the script debugger)
'   pWb->put_Silent(VARIANT_TRUE)
'   ' // Navigate to a web page
'   DIM wszUrl AS WSTRING * 260 = "https://forum.powerbasic.com/"
'   DIM vUrl AS VARIANT : vUrl.vt = VT_BSTR : vUrl.bstrVal = SysAllocString(wszUrl)
'   hr = pWb->Navigate2(@vUrl, NULL, NULL, NULL, NULL)
'   VariantClear @vurl
' ========================================================================================
PRIVATE CONSTRUCTOR CAtlCon
   CATL_DP("CAtlCon Default constructor")
   this.Init
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CAtlCon class destructor
' ========================================================================================
PRIVATE DESTRUCTOR CAtlCon
   CATL_DP("CAtlCon Destructor")
   ' // Free the ATL library
   IF m_hAtlLib THEN FreeLibrary m_hAtlLib
   ' // Free the COM library
   IF m_bUninitOLE THEN OleUninitialize
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Creates an instance of the ATL container and attaches an ActiveX control to it. 
' - pWindow: A reference to an instance of the CWindow class.
' - cID: The identifier of the control.
' - wszName: A string to be passed to the control.
'   Must be formatted in one of the following ways:
'   - A ProgID such as "MSCAL.Calendar.7"
'   - A CLSID such as "{8E27C92B-1264-101C-8A2F-040224009C02}"
'   - A URL such as "http://www.microsoft.com"
'   - A reference to an Active document such as "file://\\Documents\MyDoc.doc"
'   - A fragment of HTML such as "MSHTML:<HTML><BODY>This is a line of text</BODY></HTML>"
'     "MSHTML:" must precede the HTML fragment so that it is designated as being an MSHTML stream.
' Example:
'   ' // Create an instance of the WebBrowser control
'   DIM pCon AS CAtlCon = CAtlCon(@pWindow, IDC_WEBBROWSER, "Shell.Explorer", 0, 0, pWindow.ClientWidth, pWindow.ClientHeight)
'   ' // Make the web page themed and DPI aware
'   DIM hWb AS HWND = AtlControlWindow(pWindow.hWindow, IDC_WEBBROWSER)
'   pCon.DocHostFlags(hWb) = DOCHOSTUIFLAG_NO3DBORDER OR DOCHOSTUIFLAG_THEME OR DOCHOSTUIFLAG_DPI_AWARE
'   ' // Get a direct pointer to the Afx_IWebBrowser2 interface
'   DIM pWb AS Afx_IWebBrowser2 PTR = AtlBrowserDispatch(hWb)
'   ' // Make navigation silent (disables annoying popups of the script debugger)
'   pWb->put_Silent(VARIANT_TRUE)
'   ' // Navigate to a web page
'   DIM wszUrl AS WSTRING * 260 = "https://forum.powerbasic.com/"
'   DIM vUrl AS VARIANT : vUrl.vt = VT_BSTR : vUrl.bstrVal = SysAllocString(wszUrl)
'   pWb->Navigate2(@vUrl, NULL, NULL, NULL, NULL)
'   VariantClear @vurl
' ========================================================================================
PRIVATE CONSTRUCTOR CAtlCon (BYVAL pWindow AS CWindow PTR, BYVAL cID AS INTEGER, BYREF wszName AS WSTRING = "", _
BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
BYVAL dwStyle AS DWORD = WS_VISIBLE OR WS_CHILD OR WS_TABSTOP, BYVAL dwExStyle AS DWORD = 0)
   CATL_DP("CAtlCon Constructor CWindow Ptr")
   this.Init
   pWindow->AddControl(m_wszAtlClassName, pWindow->hWindow, cID, wszName, x, y, nWidth, nHeight, dwStyle, dwExStyle)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Creates an instance of the ATL container and attaches a licensed ActiveX control to it. 
' Parameters:
' - pWindow: A reference to an instance of the CWindow class.
' - cID: The identifier of the control.
' - wszProgID
'   The ProgID or the CLSID of the object to create.
'   - A ProgID such as "MSCAL.Calendar.7"
'   - A CLSID such as "{8E27C92B-1264-101C-8A2F-040224009C02}"
' - wszLicKey: The license key.
' ========================================================================================
PRIVATE CONSTRUCTOR CAtlCon (BYVAL pWindow AS CWindow PTR, BYVAL cID AS INTEGER, BYREF wszProgID AS WSTRING, _
BYREF wszLicKey AS WSTRING, BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
BYVAL dwStyle AS DWORD = WS_VISIBLE OR WS_CHILD OR WS_TABSTOP, BYVAL dwExStyle AS DWORD = 0)
   CATL_DP("CAtlCon Constructor ProgID/LicKey")
   this.Init
   ' // Create an instance of the licensed ActiveX control
   DIM pObj AS ANY PTR = AfxNewCom(wszProgID, wszLicKey)
   ' // Add an AtlAxWindow
   DIM hCtl AS HWND = this.AddControl(pWindow, cID, "", x, y, nWidth, nHeight, dwStyle, dwExStyle)
   ' // Attach the ActiveX control to the ATL window
   DIM hr AS HRESULT = this.AttachControl(pObj, hCtl)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Creates an instance of the ATL container and attaches a licensed ActiveX control to it. 
' Parameters:
' - pWindow: A reference to an instance of the CWindow class.
' - cID: The identifier of the control.
' - wszLibName = Full path where the library is located.
' - rclsid = The CLSID (class identifier) associated with the data and code that will be
'   used to create the object.
' - riid = A reference to the identifier of the interface to be used to communicate with the object.
' - wszLicKey = The license key.
' Not every component is a suitable candidate for use under this overloaded AfxNewCom function.
'  - Only in-process servers (DLLs) are supported.
'  - Components that are system components or part of the operating system, such as XML,
'    Data Access, Internet Explorer, or DirectX, aren't supported
'  - Components that are part of an application, such Microsoft Office, aren't supported.
'  - Components intended for use as an add-in or a snap-in, such as an Office add-in or
'    a control in a Web browser, aren't supported.
'  - Components that manage a shared physical or virtual system resource aren't supported.
' ========================================================================================
PRIVATE CONSTRUCTOR CAtlCon (BYVAL pWindow AS CWindow PTR, BYVAL cID AS INTEGER, BYREF wszLibName AS CONST WSTRING, _
BYREF rclsid AS CONST CLSID, BYREF riid AS CONST IID, BYREF wszLicKey AS WSTRING = "", _
BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
BYVAL dwStyle AS DWORD = WS_VISIBLE OR WS_CHILD OR WS_TABSTOP, BYVAL dwExStyle AS DWORD = 0)
   CATL_DP("CAtlCon Constructor LibName")
   this.Init
   ' // Create an instance of the licensed ActiveX control
   DIM pObj AS ANY PTR = AfxNewCom(wszLibName, rclsid, riid)
   ' // Add an AtlAxWindow
   DIM hCtl AS HWND = this.AddControl(pWindow, cID, "", x, y, nWidth, nHeight, dwStyle, dwExStyle)
   ' // Attach the ActiveX control to the ATL window
   DIM hr AS HRESULT = this.AttachControl(pObj, hCtl)
END CONSTRUCTOR
' ========================================================================================


' ========================================================================================
' Adds an ActiveX control to the specified window.
' Parameters:
' - pWindow: A reference to an instance of the CWindow class.
' - cID: The identifier of the control.
' - wszName: A string to be passed to the control.
'   Must be formatted in one of the following ways:
'   - A ProgID such as "MSCAL.Calendar.7"
'   - A CLSID such as "{8E27C92B-1264-101C-8A2F-040224009C02}"
'   - A URL such as "http://www.microsoft.com"
'   - A reference to an Active document such as "file://\\Documents\MyDoc.doc"
'   - A fragment of HTML such as "MSHTML:<HTML><BODY>This is a line of text</BODY></HTML>"
'     "MSHTML:" must precede the HTML fragment so that it is designated as being an MSHTML stream.
' Return value:
'   The handle of the ATL window that hosts the ActiveX control.
' Example:
'   // Create an instance of the container
'   DIM pAtlCon AS CAtlCon = CAtlCon(@pWindow, IDC_ATLWIN, "http://www.planetsquires.com/protect/forum/index.php", 10, 10, 500, 300)
'   // Get the handle of the control
'   DIM hWb AS HWND = GetDlgItem(pWindow.hWindow, IDC_ATLWIN)
'   // Make the web page themed and DPI aware
'   pAtlCon.DocHostFlags(hWb) = DOCHOSTUIFLAG_NO3DBORDER OR DOCHOSTUIFLAG_THEME OR DOCHOSTUIFLAG_DPI_AWARE
' ========================================================================================
PRIVATE FUNCTION CAtlCon.AddControl (BYVAL pWindow AS CWindow PTR, BYVAL cID AS INTEGER, _
BYREF wszName AS WSTRING = "", BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, _
BYVAL nHeight AS LONG = 0, BYVAL dwStyle AS DWORD = WS_VISIBLE OR WS_CHILD OR WS_TABSTOP, BYVAL dwExStyle AS DWORD = 0) AS HWND
   CATL_DP("CAtlCon AddControl CWindow Ptr")
   IF pWindow = NULL THEN RETURN NULL
   RETURN pWindow->AddControl(m_wszAtlClassName, pWindow->hWindow, cID, wszName, x, y, nWidth, nHeight, dwStyle, dwExStyle)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds a licensed ActiveX control to the specified window.
' Parameters:
' - pWindow: A reference to an instance of the CWindow class.
' - cID: The identifier of the control.
' - wszProgID: A string to be passed to the control.
'   Must be formatted in one of the following ways:
'   - A ProgID such as "MSCAL.Calendar.7"
'   - A CLSID such as "{8E27C92B-1264-101C-8A2F-040224009C02}"
' Return value:
'   The handle of the ATL window that hosts the ActiveX control.
' ========================================================================================
PRIVATE FUNCTION CAtlCon.AddControl (BYVAL pWindow AS CWindow PTR, BYVAL cID AS INTEGER, BYREF wszProgID AS WSTRING, _
BYREF wszLicKey AS WSTRING, BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
BYVAL dwStyle AS DWORD = WS_VISIBLE OR WS_CHILD OR WS_TABSTOP, BYVAL dwExStyle AS DWORD = 0) AS HWND
   CATL_DP("CAtlCon AddControl ProgID/LicKey")
   ' // Create an instance of the licensed ActiveX control
   DIM pObj AS ANY PTR = AfxNewCom(wszProgID, wszLicKey)
   ' // Add an AtlAxWindow
   DIM hCtl AS HWND = this.AddControl(pWindow, cID, "", x, y, nWidth, nHeight, dwStyle, dwExStyle)
   ' // Attach the ActiveX control to the ATL window
   DIM hr AS HRESULT = this.AttachControl(pObj, hCtl)
   ' // Return the handle of hosting ATL window
   RETURN hCtl
END FUNCTION
' ========================================================================================
' ========================================================================================
' Creates an instance of the ATL container and attaches a licensed ActiveX control to it. 
' Parameters:
' - pWindow: A reference to an instance of the CWindow class.
' - cID: The identifier of the control.
' - wszLibName = Full path where the library is located.
' - rclsid = The CLSID (class identifier) associated with the data and code that will be
'   used to create the object.
' - riid = A reference to the identifier of the interface to be used to communicate with the object.
' - wszLicKey = The license key.
' Not every component is a suitable candidate for use under this overloaded AfxNewCom function.
'  - Only in-process servers (DLLs) are supported.
'  - Components that are system components or part of the operating system, such as XML,
'    Data Access, Internet Explorer, or DirectX, aren't supported
'  - Components that are part of an application, such Microsoft Office, aren't supported.
'  - Components intended for use as an add-in or a snap-in, such as an Office add-in or
'    a control in a Web browser, aren't supported.
'  - Components that manage a shared physical or virtual system resource aren't supported.
' Return value:
'   The handle of the ATL window that hosts the ActiveX control.
' ========================================================================================
PRIVATE FUNCTION CAtlCon.AddControl (BYVAL pWindow AS CWindow PTR, BYVAL cID AS INTEGER, BYREF wszLibName AS CONST WSTRING, _
BYREF rclsid AS CONST CLSID, BYREF riid AS CONST IID, BYREF wszLicKey AS WSTRING = "", _
BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
BYVAL dwStyle AS DWORD = WS_VISIBLE OR WS_CHILD OR WS_TABSTOP, BYVAL dwExStyle AS DWORD = 0) AS HWND
   CATL_DP("CAtlCon AddControl LibName")
   ' // Create an instance of the licensed ActiveX control
   DIM pObj AS ANY PTR = AfxNewCom(wszLibName, rclsid, riid)
   ' // Add an AtlAxWindow
   DIM hCtl AS HWND = this.AddControl(pWindow, cID, "", x, y, nWidth, nHeight, dwStyle, dwExStyle)
   ' // Attach the ActiveX control to the ATL window
   DIM hr AS HRESULT = this.AttachControl(pObj, hCtl)
   ' // Return the handle of hosting ATL window
   RETURN hCtl
END FUNCTION
' ========================================================================================

' ========================================================================================
' Attaches a previously created ActiveX control to the specified window.
' Parameters:
' - pControl: A pointer to the IUnknown of the control.
' - hCtl: Handle to the window that will host the control.
' ========================================================================================
PRIVATE FUNCTION CAtlCon.AttachControl (BYVAL pControl AS IUnknown PTR, BYVAL hCtl AS HWND) AS HRESULT
   CATL_DP("CAtlCon AttachControl")
   DIM pProc AS FUNCTION (BYVAL pControl AS IUnknown PTR, BYVAL hCtl AS HWND, BYVAL ppUnkContainer AS IUnknown PTR PTR) AS HRESULT
   pProc = CAST(ANY PTR, GetProcAddress(m_hAtlLib, "AtlAxAttachControl"))
   IF pProc = NULL THEN RETURN FALSE
   RETURN pProc(pControl, hCtl, NULL)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the version of the ATL library that you are using.
' ========================================================================================
PRIVATE FUNCTION CAtlCon.AtlVersion () AS DWORD
   CATL_DP("CAtlCon AtlVersion")
   DIM pProc AS FUNCTION () AS DWORD
   pProc = CAST(ANY PTR, GetProcAddress(m_hAtlLib, "AtlVersion"))
   IF pProc = NULL THEN RETURN 0
   RETURN pProc()
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts an object's size in HIMETRIC units (each unit is 0.01 millimeter) to a size in
' pixels on the screen device.
' - SizeInHiMetric: Size of the object in HIMETRIC units
' ========================================================================================
PRIVATE FUNCTION CAtlCon.HiMetricToPixel (BYREF SizeInHiMetric AS SIZEL) AS SIZEL
   CATL_DP("CAtlCon HiMetricToPixel")
   DIM pProc AS SUB (BYREF SizeInHiMetric AS SIZEL, BYREF SizeInPix AS SIZEL)
   pProc = CAST(ANY PTR, GetProcAddress(m_hAtlLib, "AtlHiMetricToPixel"))
   DIM SizeInPix AS SIZEL
   IF pProc = NULL THEN pProc(SizeInHiMetric, SizeInPix)
   RETURN SizeInPix
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts an object's size in pixels on the screen device to a size in HIMETRIC units
' (each unit is 0.01 millimeter).
' - SizeInHiMetric: Size of the object in HIMETRIC units
' ========================================================================================
PRIVATE FUNCTION CAtlCon.PixelToHiMetric (BYREF SizeInPix AS SIZEL) AS SIZEL
   CATL_DP("CAtlCon PixelToHiMetric")
   DIM pProc AS SUB (BYREF SizeInPix AS SIZEL, BYREF SizeInHiMetric AS SIZEL)
   pProc = CAST(ANY PTR, GetProcAddress(m_hAtlLib, "AtlPixelToHiMetric"))
   DIM SizeInHiMetric AS SIZEL
   IF pProc = NULL THEN pProc(SizeInPix, SizeInHiMetric)
   RETURN SizeInHiMetric
END FUNCTION
' ========================================================================================

' ========================================================================================
' Obtains a direct interface pointer to the container for a specified window (if any),
' given its handle.
' - hwnd:  A handle to the window that is hosting the control.
' Return value:
'   The IUnknown of the container of the control.
' ========================================================================================
PRIVATE FUNCTION CAtlCon.GetHost (BYVAL hwnd AS .HWND) AS IUnknown PTR
   CATL_DP("CAtlCon GetHost")
   DIM pProc AS FUNCTION (BYVAL hwnd AS .HWND, BYREF punk AS IUnknown PTR) AS HRESULT
   pProc = CAST(ANY PTR, GetProcAddress(m_hAtlLib, "AtlAxGetHost"))
   IF pProc = NULL THEN RETURN NULL
   DIM punk AS IUnknown PTR
   DIM hr AS HRESULT = pProc(hwnd, punk)
   RETURN punk
END FUNCTION
' ========================================================================================

' ========================================================================================
' Obtains a direct interface pointer to the control contained inside a specified window
' given its handle. Don't release this pointer because it is not an addrefed pointer.
' ========================================================================================
PRIVATE FUNCTION CAtlCon.GetControl (BYVAL hwnd AS .HWND) AS IUnknown PTR
   CATL_DP("CAtlCon GetControl")
   DIM pProc AS FUNCTION (BYVAL hwnd AS .HWND, BYREF punk AS IUnknown PTR) AS HRESULT
   pProc = CAST(ANY PTR, GetProcAddress(m_hAtlLib, "AtlAxGetControl"))
   IF pProc = NULL THEN RETURN NULL
   DIM punk AS IUnknown PTR
   DIM hr AS HRESULT = pProc(hwnd, punk)
   RETURN punk
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the dispatch interface of the ActiveX control given the handle of its ATL container.
' Release this pointer when no longer needed.
' ========================================================================================
PRIVATE FUNCTION CAtlCon.GetDispatch (BYVAL hwnd AS .HWND) AS IDispatch PTR
   CATL_DP("CAtlCon GetDispatch")
   ' // Get the IUnknown of the OCX hosted in the control
   DIM punk AS IUnknown PTR = this.GetControl(hwnd)
   IF punk = NULL THEN RETURN NULL
   ' // Query for the IDispatch interface
   DIM pdisp AS IDispatch PTR
   pUnk->lpVtbl->QueryInterface(pUnk, @IID_IDispatch, @pdisp)
   RETURN pdisp
END FUNCTION
' ========================================================================================

' ========================================================================================
' Specifies the user interface capabilities of the host object.
' DOCHOSTUIFLAG_DIALOG
'     MSHTML does not enable selection of the text in the form.
' DOCHOSTUIFLAG_DISABLE_HELP_MENU
'     MSHTML does not add the Help menu item to the container's menu.
' DOCHOSTUIFLAG_NO3DBORDER
'     MSHTML does not use 3-D borders on any frames or framesets. To turn the border off
'     on only the outer frameset use DOCHOSTUIFLAG_NO3DOUTERBORDER
' DOCHOSTUIFLAG_SCROLL_NO
'     MSHTML does not have scroll bars.
' DOCHOSTUIFLAG_DISABLE_SCRIPT_INACTIVE
'     MSHTML does not execute any script until fully activated. This flag is used to
'     postpone script execution until the host is active and, therefore, ready for script
'     to be executed.
' DOCHOSTUIFLAG_OPENNEWWIN
'     MSHTML opens a site in a new window when a link is clicked rather than browse to
'     the new site using the same browser window.
' DOCHOSTUIFLAG_DISABLE_OFFSCREEN
'     Not implemented.
' DOCHOSTUIFLAG_FLAT_SCROLLBAR
'     MSHTML uses flat scroll bars for any UI it displays.
' DOCHOSTUIFLAG_DIV_BLOCKDEFAULT
'     MSHTML inserts the div tag if a return is entered in edit mode. Without this flag,
'     MSHTML will use the p tag.
' DOCHOSTUIFLAG_ACTIVATE_CLIENTHIT_ONLY
'     MSHTML only becomes UI active if the mouse is clicked in the client area of the
'     window. It does not become UI active if the mouse is clicked on a non-client area,
'     such as a scroll bar.
' DOCHOSTUIFLAG_OVERRIDEBEHAVIORFACTORY
'     MSHTML consults the host before retrieving a behavior from the URL specified on
'     the page. If the host does not support the behavior, MSHTML does not proceed to
'     query other hosts or instantiate the behavior itself, even for behaviors developed
'     in script (HTML Components (HTCs)).
' DOCHOSTUIFLAG_CODEPAGELINKEDFONTS
'     Microsoft Internet Explorer 5 and later. Provides font selection compatibility
'     for Microsoft Outlook Express. If the flag is enabled, the displayed characters
'     are inspected to determine whether the current font supports the code page. If
'     disabled, the current font is used, even if it does not contain a glyph for the
'     character. This flag assumes that the user is using Internet Explorer 5 and
'     Outlook Express 4.0.
' DOCHOSTUIFLAG_URL_ENCODING_DISABLE_UTF8
'     Internet Explorer 5 and later. Controls how nonnative URLs are transmitted over
'     the Internet. Nonnative refers to characters outside the multibyte encoding of
'     the URL. If this flag is set, the URL is not submitted to the server in UTF-8 encoding.
' DOCHOSTUIFLAG_URL_ENCODING_ENABLE_UTF8
'     Internet Explorer 5 and later. Controls how nonnative URLs are transmitted over
'     the Internet. Nonnative refers to characters outside the multibyte encoding of
'     the URL. If this flag is set, the URL is submitted to the server in UTF-8 encoding.
' DOCHOSTUIFLAG_ENABLE_FORMS_AUTOCOMPLETE
'     Internet Explorer 5 and later. Enables the AutoComplete feature for forms in the
'     hosted browser. The Intelliforms feature is only turned on if the user has
'     previously enabled it. If the user has turned the AutoComplete feature off for
'     forms, it is off whether this flag is specified or not.
' DOCHOSTUIFLAG_ENABLE_INPLACE_NAVIGATION
'     Internet Explorer 5 and later. Enables the host to specify that navigation should
'     happen in place. This means that applications hosting MSHTML directly can specify
'     that navigation happen in the application's window. For instance, if this flag is
'     set, you can click a link in HTML mail and navigate in the mail instead of opening
'     a new Windows Internet Explorer window.
' DOCHOSTUIFLAG_IME_ENABLE_RECONVERSION
'     Internet Explorer 5 and later. During initialization, the host can set this flag
'     to enable Input Method Editor (IME) reconversion, allowing computer users to employ
'     IME reconversion while browsing Web pages. An input method editor is a program that
'     allows users to enter complex characters and symbols, such as Japanese Kanji
'     characters, using a standard keyboard. For more information, see the International
'     Features reference in the Base Services section of the Windows Software Development
'     Kit (SDK).
' DOCHOSTUIFLAG_THEME
'     Internet Explorer 6 and later. Specifies that the hosted browser should use themes
'     for pages it displays.
' DOCHOSTUIFLAG_NOTHEME
'     Internet Explorer 6 and later. Specifies that the hosted browser should not use
'     themes for pages it displays.
' DOCHOSTUIFLAG_NOPICS
'     Internet Explorer 6 and later. Disables PICS ratings for the hosted browser.
' DOCHOSTUIFLAG_NO3DOUTERBORDER
'     Internet Explorer 6 and later. Turns off any 3-D border on the outermost frame or
'     frameset only. To turn borders off on all frame sets, use DOCHOSTUIFLAG_NO3DBORDER
' DOCHOSTUIFLAG_DISABLE_EDIT_NS_FIXUP
'     Internet Explorer 6 and later. Disables the automatic correction of namespaces when
'     editing HTML elements.
' DOCHOSTUIFLAG_LOCAL_MACHINE_ACCESS_CHECK
'     Internet Explorer 6 and later. Prevents Web sites in the Internet zone from accessing
'     files in the Local Machine zone.
' DOCHOSTUIFLAG_DISABLE_UNTRUSTEDPROTOCOL
'     Internet Explorer 6 and later. Turns off untrusted protocols. Untrusted protocols
'     include ms-its, ms-itss, its, and mk:@msitstore.
' DOCHOSTUIFLAG_HOST_NAVIGATES
'     Internet Explorer 7. Indicates that navigation is delegated to the host; otherwise,
'     MSHTML will perform navigation. This flag is used primarily for non-HTML document types.
' DOCHOSTUIFLAG_ENABLE_REDIRECT_NOTIFICATION
'     Internet Explorer 7. Causes MSHTML to fire an additional DWebBrowserEvents2::BeforeNavigate2
'     event when redirect navigations occur. Applications hosting the WebBrowser Control
'     can choose to cancel or continue the redirect by returning an appropriate value in
'     the Cancel parameter of the event.
' DOCHOSTUIFLAG_USE_WINDOWLESS_SELECTCONTROL
'     Internet Explorer 7. Causes MSHTML to use the Document Object Model (DOM) to create
'     native "windowless" select controls that can be visually layered under other elements.
' DOCHOSTUIFLAG_USE_WINDOWED_SELECTCONTROL
'     Internet Explorer 7. Causes MSHTML to create standard Microsoft Win32 "windowed"
'     select and drop-down controls.
' DOCHOSTUIFLAG_ENABLE_ACTIVEX_INACTIVATE_MODE
'     Internet Explorer 6 for Windows XP Service Pack 2 (SP2) and later. Requires user
'     activation for Microsoft ActiveX controls and Java Applets embedded within a web page.
'     This flag enables interactive control blocking, which provisionally disallows direct
'     interaction with ActiveX controls loaded by the APPLET, EMBED, or OBJECT elements.
'     When a control is inactive, it does not respond to user input; however, it can perform
'     operations that do not involve interaction.
' DOCHOSTUIFLAG_DPI_AWARE
'     Internet Explorer 8. Causes layout engine to calculate document pixels as 96 dots
'     per inch (dpi). Normally, a document pixel is the same size as a screen pixel. This
'     flag is equivalent to setting the FEATURE_96DPI_PIXEL feature control key on a
'     per-host basis.
' Remarks
'     The DOCHOSTUIFLAG_BROWSER flag, a supplementary defined constant (not technically
'     a part of this enumeration), combines the values of DOCHOSTUIFLAG_DISABLE_HELP_MENU
'     and DOCHOSTUIFLAG_DISABLE_SCRIPT_INACTIVE.
'     The ATL host object implementation uses DOCHOSTUIFLAG_NO3DBORDER as the default
'     value of this property.
' Note: Don't release pDocHost because it is a direct pointer.
' Example: pAtlCon.DocHostFlags(hWb) = DOCHOSTUIFLAG_NO3DBORDER OR DOCHOSTUIFLAG_THEME OR DOCHOSTUIFLAG_DPI_AWARE
' ========================================================================================
PRIVATE PROPERTY CAtlCon.DocHostFlags (BYVAL hwnd AS .HWND, BYVAL dwDocHostFlags AS DWORD)
   CATL_DP("CAtlCon put_DocHostFlags")
   DIM pDocHost AS IUnknown PTR = this.GetHost(hwnd)
   DIM pAmbientDisp AS Afx_IAxWinAmbientDispatch PTR
   IF pDocHost THEN
      pDocHost->lpvtbl->QueryInterface(pDocHost, @AFX_IID_IAxWinAmbientDispatch_, @pAmbientDisp)
      IF pAmbientDisp THEN
         pAmbientDisp->put_DocHostFlags(DOCHOSTUIFLAG_NO3DBORDER OR DOCHOSTUIFLAG_THEME OR DOCHOSTUIFLAG_DPI_AWARE)
         pAmbientDisp->Release
      END IF
   END IF
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CAtlCon.DocHostFlags (BYVAL hwnd AS .HWND) AS DWORD
   CATL_DP("CAtlCon get_DocHostFlags")
   DIM pDocHost AS IUnknown PTR = this.GetHost(hwnd)
   DIM pAmbientDisp AS Afx_IAxWinAmbientDispatch PTR
   DIM dwDocHostFlags AS DWORD
   IF pDocHost THEN
      pDocHost->lpvtbl->QueryInterface(pDocHost, @AFX_IID_IAxWinAmbientDispatch_, @pAmbientDisp)
      IF pAmbientDisp THEN
         pAmbientDisp->get_DocHostFlags(@dwDocHostFlags)
         pAmbientDisp->Release
      END IF
   END IF
   PROPERTY = dwDocHostFlags
END PROPERTY
' ========================================================================================

' ========================================================================================
' Specifies the operation that should take place in response to a double-click.
' - dwDocHostDoubleClickFlags: The new value of this property.
' Constants:
' DOCHOSTUIDBLCLK_DEFAULT
'    Perform the default action.
' DOCHOSTUIDBLCLK_SHOWPROPERTIES
'    Show the item's properties.
' DOCHOSTUIDBLCLK_SHOWCODE
'    Show the page's source.
' The ATL host object implementation uses DOCHOSTUIDBLCLK_DEFAULT as the default value
' of this property.
' ========================================================================================
PRIVATE PROPERTY CAtlCon.DocHostDoubleClickFlags (BYVAL hwnd AS .HWND, BYVAL dwDocHostDoubleClickFlags AS DWORD)
   CATL_DP("CAtlCon put_DocHostDoubleClickFlags")
   DIM pDocHost AS IUnknown PTR = this.GetHost(hwnd)
   DIM pAmbientDisp AS Afx_IAxWinAmbientDispatch PTR
   IF pDocHost THEN
      pDocHost->lpvtbl->QueryInterface(pDocHost, @AFX_IID_IAxWinAmbientDispatch_, @pAmbientDisp)
      IF pAmbientDisp THEN
         pAmbientDisp->put_DocHostDoubleClickFlags(dwDocHostDoubleClickFlags)
         pAmbientDisp->Release
      END IF
   END IF
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CAtlCon.DocHostDoubleClickFlags (BYVAL hwnd AS .HWND) AS DWORD
   CATL_DP("CAtlCon get_DocHostDoubleClickFlags")
   DIM pDocHost AS IUnknown PTR = this.GetHost(hwnd)
   DIM pAmbientDisp AS Afx_IAxWinAmbientDispatch PTR
   DIM dwDocHostDoubleClickFlags AS DWORD
   IF pDocHost THEN
      pDocHost->lpvtbl->QueryInterface(pDocHost, @AFX_IID_IAxWinAmbientDispatch_, @pAmbientDisp)
      IF pAmbientDisp THEN
         pAmbientDisp->get_DocHostDoubleClickFlags(@dwDocHostDoubleClickFlags)
         pAmbientDisp->Release
      END IF
   END IF
   PROPERTY = dwDocHostDoubleClickFlags
END PROPERTY
' ========================================================================================

' ========================================================================================
' Specifies whether the hosted control is allowed to display its own context menu.
' - bAllowContextMenu: VARIANT_TRUE (context menu allowed), VARIANT_FALSE (context menu disabled).
' The ATL host object implementation uses VARIANT_TRUE as the default value of this property.
' ========================================================================================
PRIVATE PROPERTY CAtlCon.AllowContextMenu (BYVAL hwnd AS .HWND, BYVAL bAllowContextMenu AS VARIANT_BOOL)
   CATL_DP("CAtlCon put_AllowContextMenu")
   DIM pDocHost AS IUnknown PTR = this.GetHost(hwnd)
   DIM pAmbientDisp AS Afx_IAxWinAmbientDispatch PTR
   IF pDocHost THEN
      pDocHost->lpvtbl->QueryInterface(pDocHost, @AFX_IID_IAxWinAmbientDispatch_, @pAmbientDisp)
      IF pAmbientDisp THEN
         pAmbientDisp->put_AllowContextMenu(bAllowContextMenu)
         pAmbientDisp->Release
      END IF
   END IF
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CAtlCon.AllowContextMenu (BYVAL hwnd AS .HWND) AS VARIANT_BOOL
   CATL_DP("CAtlCon get_AllowContextMenu")
   DIM pDocHost AS IUnknown PTR = this.GetHost(hwnd)
   DIM pAmbientDisp AS Afx_IAxWinAmbientDispatch PTR
   DIM bAllowContextMenu AS VARIANT_BOOL
   IF pDocHost THEN
      pDocHost->lpvtbl->QueryInterface(pDocHost, @AFX_IID_IAxWinAmbientDispatch_, @pAmbientDisp)
      IF pAmbientDisp THEN
         pAmbientDisp->get_AllowContextMenu(@bAllowContextMenu)
         pAmbientDisp->Release
      END IF
   END IF
   PROPERTY = bAllowContextMenu
END PROPERTY
' ========================================================================================

' ========================================================================================
' Specifies whether the hosted control is allowed to display its own user interface.
' - bAllowShowUI: VARIANT_TRUE or VARIANT_FALSE.
' The ATL host object implementation uses VARIANT_TRUE as the default value of this property.
' ========================================================================================
PRIVATE PROPERTY CAtlCon.AllowShowUI (BYVAL hwnd AS .HWND, BYVAL bAllowShowUI AS VARIANT_BOOL)
   CATL_DP("CAtlCon put_AllowShowUI")
   DIM pDocHost AS IUnknown PTR = this.GetHost(hwnd)
   DIM pAmbientDisp AS Afx_IAxWinAmbientDispatch PTR
   IF pDocHost THEN
      pDocHost->lpvtbl->QueryInterface(pDocHost, @AFX_IID_IAxWinAmbientDispatch_, @pAmbientDisp)
      IF pAmbientDisp THEN
         pAmbientDisp->put_AllowShowUI(bAllowShowUI)
         pAmbientDisp->Release
      END IF
   END IF
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CAtlCon.AllowShowUI (BYVAL hwnd AS .HWND) AS VARIANT_BOOL
   CATL_DP("CAtlCon put_AllowShowUI")
   DIM pDocHost AS IUnknown PTR = this.GetHost(hwnd)
   DIM pAmbientDisp AS Afx_IAxWinAmbientDispatch PTR
   DIM bAllowShowUI AS VARIANT_BOOL
   IF pDocHost THEN
      pDocHost->lpvtbl->QueryInterface(pDocHost, @AFX_IID_IAxWinAmbientDispatch_, @pAmbientDisp)
      IF pAmbientDisp THEN
         pAmbientDisp->get_AllowShowUI(@bAllowShowUI)
         pAmbientDisp->Release
      END IF
   END IF
   PROPERTY = bAllowShowUI
END PROPERTY
' ========================================================================================

' ========================================================================================
' Specifies whether the container will allow windowless activation.
' - bAllowWindowless: VARIANT_TRUE or VARIANT_FALSE.
' The ATL host object implementation uses VARIANT_TRUE as the default value of this property.
' ========================================================================================
PRIVATE PROPERTY CAtlCon.AllowWindowlessActivation (BYVAL hwnd AS .HWND, BYVAL bAllowWindowless AS VARIANT_BOOL)
   CATL_DP("CAtlCon put_AllowWindowlessActivation")
   DIM pDocHost AS IUnknown PTR = this.GetHost(hwnd)
   DIM pAmbientDisp AS Afx_IAxWinAmbientDispatch PTR
   IF pDocHost THEN
      pDocHost->lpvtbl->QueryInterface(pDocHost, @AFX_IID_IAxWinAmbientDispatch_, @pAmbientDisp)
      IF pAmbientDisp THEN
         pAmbientDisp->put_AllowWindowlessActivation(bAllowWindowless)
         pAmbientDisp->Release
      END IF
   END IF
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CAtlCon.AllowWindowlessActivation (BYVAL hwnd AS .HWND) AS VARIANT_BOOL
   CATL_DP("CAtlCon get_AllowWindowlessActivation")
   DIM pDocHost AS IUnknown PTR = this.GetHost(hwnd)
   DIM pAmbientDisp AS Afx_IAxWinAmbientDispatch PTR
   DIM bAllowWindowless AS VARIANT_BOOL
   IF pDocHost THEN
      pDocHost->lpvtbl->QueryInterface(pDocHost, @AFX_IID_IAxWinAmbientDispatch_, @pAmbientDisp)
      IF pAmbientDisp THEN
         pAmbientDisp->get_AllowWindowlessActivation(@bAllowWindowless)
         pAmbientDisp->Release
      END IF
   END IF
   PROPERTY = bAllowWindowless
END PROPERTY
' ========================================================================================

' ========================================================================================
' Specifies the ambient background color of the container.
' - clrBackground: The new value of this property.
' The ATL host object implementation uses COLOR_BTNFACE or COLOR_WINDOW as the default value
' of this property (depending on whether the parent of the host window is a dialog or not).
' ========================================================================================
PRIVATE PROPERTY CAtlCon.BackColor (BYVAL hwnd AS .HWND, BYVAL clrBackground AS OLE_COLOR)
   CATL_DP("CAtlCon put_BackColor")
   DIM pDocHost AS IUnknown PTR = this.GetHost(hwnd)
   DIM pAmbientDisp AS Afx_IAxWinAmbientDispatch PTR
   IF pDocHost THEN
      pDocHost->lpvtbl->QueryInterface(pDocHost, @AFX_IID_IAxWinAmbientDispatch_, @pAmbientDisp)
      IF pAmbientDisp THEN
         pAmbientDisp->put_BackColor(clrBackground)
         pAmbientDisp->Release
      END IF
   END IF
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CAtlCon.BackColor (BYVAL hwnd AS .HWND) AS OLE_COLOR
   CATL_DP("CAtlCon get_BackColor")
   DIM pDocHost AS IUnknown PTR = this.GetHost(hwnd)
   DIM pAmbientDisp AS Afx_IAxWinAmbientDispatch PTR
   DIM clrBackground AS OLE_COLOR
   IF pDocHost THEN
      pDocHost->lpvtbl->QueryInterface(pDocHost, @AFX_IID_IAxWinAmbientDispatch_, @pAmbientDisp)
      IF pAmbientDisp THEN
         pAmbientDisp->get_BackColor(@clrBackground)
         pAmbientDisp->Release
      END IF
   END IF
   PROPERTY = clrBackground
END PROPERTY
' ========================================================================================

' ========================================================================================
' Specifies the ambient foreground color of the container.
' - clrForeground: The new value of this property.
' The ATL host object implementation uses the system window text color as the default
' value of this property.
' ========================================================================================
PRIVATE PROPERTY CAtlCon.ForeColor (BYVAL hwnd AS .HWND, BYVAL clrForeground AS OLE_COLOR)
   CATL_DP("CAtlCon put_ForeColor")
   DIM pDocHost AS IUnknown PTR = this.GetHost(hwnd)
   DIM pAmbientDisp AS Afx_IAxWinAmbientDispatch PTR
   IF pDocHost THEN
      pDocHost->lpvtbl->QueryInterface(pDocHost, @AFX_IID_IAxWinAmbientDispatch_, @pAmbientDisp)
      IF pAmbientDisp THEN
         pAmbientDisp->put_ForeColor(clrForeground)
         pAmbientDisp->Release
      END IF
   END IF
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CAtlCon.ForeColor (BYVAL hwnd AS .HWND) AS OLE_COLOR
   CATL_DP("CAtlCon get_ForeColor")
   DIM pDocHost AS IUnknown PTR = this.GetHost(hwnd)
   DIM pAmbientDisp AS Afx_IAxWinAmbientDispatch PTR
   DIM clrForeground AS OLE_COLOR
   IF pDocHost THEN
      pDocHost->lpvtbl->QueryInterface(pDocHost, @AFX_IID_IAxWinAmbientDispatch_, @pAmbientDisp)
      IF pAmbientDisp THEN
         pAmbientDisp->get_ForeColor(@clrForeground)
         pAmbientDisp->Release
      END IF
   END IF
   PROPERTY = clrForeground
END PROPERTY
' ========================================================================================

' ========================================================================================
' Specifies the ambient locale ID of the container.
' - lcidLocaleID: The new value of this property.
' The ATL host object implementation uses the user's default locale as the default value
' of this property.
' ========================================================================================
PRIVATE PROPERTY CAtlCon.LocaleID (BYVAL hwnd AS .HWND, BYVAL lcidLocaleID AS LCID)
   CATL_DP("CAtlCon put_LocaleID")
   DIM pDocHost AS IUnknown PTR = this.GetHost(hwnd)
   DIM pAmbientDisp AS Afx_IAxWinAmbientDispatch PTR
   IF pDocHost THEN
      pDocHost->lpvtbl->QueryInterface(pDocHost, @AFX_IID_IAxWinAmbientDispatch_, @pAmbientDisp)
      IF pAmbientDisp THEN
         pAmbientDisp->put_LocaleID(lcidLocaleID)
         pAmbientDisp->Release
      END IF
   END IF
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CAtlCon.LocaleID (BYVAL hwnd AS .HWND) AS LCID
   CATL_DP("CAtlCon get_LocaleID")
   DIM pDocHost AS IUnknown PTR = this.GetHost(hwnd)
   DIM pAmbientDisp AS Afx_IAxWinAmbientDispatch PTR
   DIM lcidLocaleID AS LCID
   IF pDocHost THEN
      pDocHost->lpvtbl->QueryInterface(pDocHost, @AFX_IID_IAxWinAmbientDispatch_, @pAmbientDisp)
      IF pAmbientDisp THEN
         pAmbientDisp->get_LocaleID(@lcidLocaleID)
         pAmbientDisp->Release
      END IF
   END IF
   PROPERTY = lcidLocaleID
END PROPERTY
' ========================================================================================

' ========================================================================================
' Ambient property that allows a control to find out if it is the default control.
' - bDisplayAsDefault: The new value of this property.
' The ATL host object implementation uses VARIANT_FALSE as the default value of this property.
' ========================================================================================
PRIVATE PROPERTY CAtlCon.DisplayAsDefault (BYVAL hwnd AS .HWND, BYVAL bDisplayAsDefault AS VARIANT_BOOL)
   CATL_DP("CAtlCon put_DisplayAsDefault")
   DIM pDocHost AS IUnknown PTR = this.GetHost(hwnd)
   DIM pAmbientDisp AS Afx_IAxWinAmbientDispatch PTR
   IF pDocHost THEN
      pDocHost->lpvtbl->QueryInterface(pDocHost, @AFX_IID_IAxWinAmbientDispatch_, @pAmbientDisp)
      IF pAmbientDisp THEN
         pAmbientDisp->put_DisplayAsDefault(bDisplayAsDefault)
         pAmbientDisp->Release
      END IF
   END IF
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CAtlCon.DisplayAsDefault (BYVAL hwnd AS .HWND) AS VARIANT_BOOL
   CATL_DP("CAtlCon get_DisplayAsDefault")
   DIM pDocHost AS IUnknown PTR = this.GetHost(hwnd)
   DIM pAmbientDisp AS Afx_IAxWinAmbientDispatch PTR
   DIM bDisplayAsDefault AS VARIANT_BOOL
   IF pDocHost THEN
      pDocHost->lpvtbl->QueryInterface(pDocHost, @AFX_IID_IAxWinAmbientDispatch_, @pAmbientDisp)
      IF pAmbientDisp THEN
         pAmbientDisp->get_DisplayAsDefault(@bDisplayAsDefault)
         pAmbientDisp->Release
      END IF
   END IF
   PROPERTY = bDisplayAsDefault
END PROPERTY
' ========================================================================================

' ========================================================================================
' Specifies whether the container will reflect messages to the hosted control.
' - bMessageReflect: VARIANT_TRUE or VARIANT_FALSE.
' The ATL host object implementation uses VARIANT_TRUE as the default value of this property.
' ========================================================================================
PRIVATE PROPERTY CAtlCon.MessageReflect (BYVAL hwnd AS .HWND, BYVAL bMessageReflect AS VARIANT_BOOL)
   CATL_DP("CAtlCon put_MessageReflect")
   DIM pDocHost AS IUnknown PTR = this.GetHost(hwnd)
   DIM pAmbientDisp AS Afx_IAxWinAmbientDispatch PTR
   IF pDocHost THEN
      pDocHost->lpvtbl->QueryInterface(pDocHost, @AFX_IID_IAxWinAmbientDispatch_, @pAmbientDisp)
      IF pAmbientDisp THEN
         pAmbientDisp->put_MessageReflect(bMessageReflect)
         pAmbientDisp->Release
      END IF
   END IF
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CAtlCon.MessageReflect (BYVAL hwnd AS .HWND) AS VARIANT_BOOL
   CATL_DP("CAtlCon get_MessageReflect")
   DIM pDocHost AS IUnknown PTR = this.GetHost(hwnd)
   DIM pAmbientDisp AS Afx_IAxWinAmbientDispatch PTR
   DIM bMessageReflect AS VARIANT_BOOL
   IF pDocHost THEN
      pDocHost->lpvtbl->QueryInterface(pDocHost, @AFX_IID_IAxWinAmbientDispatch_, @pAmbientDisp)
      IF pAmbientDisp THEN
         pAmbientDisp->get_MessageReflect(@bMessageReflect)
         pAmbientDisp->Release
      END IF
   END IF
   PROPERTY = bMessageReflect
END PROPERTY
' ========================================================================================

' ========================================================================================
' Specifies the ambient font of the container.
' - pFont: Pointer to the IFontDisp object to set.
' The ATL host object implementation uses the default GUI font or the system font as the
' default value of this property.
' ========================================================================================
PRIVATE PROPERTY CAtlCon.Font (BYVAL hwnd AS .HWND, BYVAL pFont AS IFontDisp PTR)
   CATL_DP("CAtlCon put_Font")
   DIM pDocHost AS IUnknown PTR = this.GetHost(hwnd)
   DIM pAmbientDisp AS Afx_IAxWinAmbientDispatch PTR
   IF pDocHost THEN
      pDocHost->lpvtbl->QueryInterface(pDocHost, @AFX_IID_IAxWinAmbientDispatch_, @pAmbientDisp)
      IF pAmbientDisp THEN
         pAmbientDisp->put_Font(pFont)
         pAmbientDisp->Release
      END IF
   END IF
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CAtlCon.Font (BYVAL hwnd AS .HWND) AS IFontDisp PTR
   CATL_DP("CAtlCon get_Font")
   DIM pDocHost AS IUnknown PTR = this.GetHost(hwnd)
   DIM pAmbientDisp AS Afx_IAxWinAmbientDispatch PTR
   DIM pFont AS IFontDisp PTR
   IF pDocHost THEN
      pDocHost->lpvtbl->QueryInterface(pDocHost, @AFX_IID_IAxWinAmbientDispatch_, @pAmbientDisp)
      IF pAmbientDisp THEN
         pAmbientDisp->get_Font(@pFont)
         pAmbientDisp->Release
      END IF
   END IF
   PROPERTY = pFont
END PROPERTY
' ========================================================================================

' ========================================================================================
' Creates a connection between an object's connection point and a client's sink.
' - pUnkCP: [in] A pointer to the IUnknown of the object that the client is connected with.
' - pUnk: [in] A pointer to the client's IUnknown.
' - iid: [in] The GUID of the connection point. Typically, this is the same as the outgoing
'   interface managed by the connection point.
' - dwCookie: [out] A cookie that uniquely identifies the connection.
' ========================================================================================
PRIVATE FUNCTION CAtlCon.Advise (BYVAL punkCP AS IUnknown PTR, BYVAL punk AS IUnknown PTR, BYVAL riid AS IID PTR, BYREF dwCookie AS DWORD) AS HRESULT
   CATL_DP("CAtlCon Advise")
   DIM pProc AS FUNCTION (BYVAL punkCP AS IUnknown PTR, BYVAL punk AS IUnknown PTR, BYVAL riid AS IID PTR, BYREF dwCookie AS DWORD) AS HRESULT
   pProc = CAST(ANY PTR, GetProcAddress(m_hAtlLib, "AtlAdvise"))
   IF pProc = NULL THEN RETURN 0
   RETURN pProc(punkCP, punk, riid, dwCookie)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Terminates the connection established through Advise.
' Parameters:
' - pUnkCP: [in] A pointer to the IUnknown of the object that the client is connected with.
' - iid: [in] The GUID of the connection point. Typically, this is the same as the outgoing
'   interface managed by the connection point.
' - dwCookie: [in] The cookie that uniquely identifies the connection.
' ========================================================================================
PRIVATE FUNCTION CAtlCon.Unadvise (BYVAL punkCP AS IUnknown PTR, BYVAL riid AS IID PTR, BYVAL dwCookie AS DWORD) AS HRESULT
   CATL_DP("CAtlCon Unadvise")
   DIM pProc AS FUNCTION (BYVAL punkCP AS IUnknown PTR, BYVAL riid AS IID PTR, BYVAL dwCookie AS DWORD) AS HRESULT
   pProc = CAST(ANY PTR, GetProcAddress(m_hAtlLib, "AtlUnadvise"))
   IF pProc = NULL THEN RETURN 0
   RETURN pProc(punkCP, riid, dwCookie)
END FUNCTION
' ========================================================================================

' ########################################################################################
'                                *** CWebCtx class ***
' ########################################################################################

' ========================================================================================
' CWebCtx class
' Implements a WebBrowser clas using ATL as its OLE container.
' ========================================================================================
TYPE CWebCtx EXTENDS CAtlCon

Private:
   m_hWb AS HWND                   ' // Handle of the ATL control window
   m_pWb AS Afx_IWebBrowser2 PTR   ' // Pointer to the IWebBrowser2 interface
   m_dwCookie AS DWORD             ' // Event's connection cookies

Public:
   DECLARE CONSTRUCTOR (BYREF pWindow AS CWindow, BYVAL cID AS INTEGER, BYREF wszURL AS WSTRING = "about:blank", _
           BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
           BYVAL dwStyle AS DWORD = WS_VISIBLE OR WS_CHILD OR WS_TABSTOP, BYVAL dwExStyle AS DWORD = 0, _
           BYVAL bSilent AS VARIANT_BOOL = VARIANT_TRUE, BYVAL dwDocHostFlags AS DWORD = DOCHOSTUIFLAG_NO3DBORDER OR DOCHOSTUIFLAG_THEME OR DOCHOSTUIFLAG_DPI_AWARE)
   DECLARE CONSTRUCTOR (BYVAL pWindow AS CWindow PTR, BYVAL cID AS INTEGER, BYREF wszURL AS WSTRING = "about:blank", _
           BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
           BYVAL dwStyle AS DWORD = WS_VISIBLE OR WS_CHILD OR WS_TABSTOP, BYVAL dwExStyle AS DWORD = 0, _
           BYVAL bSilent AS VARIANT_BOOL = VARIANT_TRUE, BYVAL dwDocHostFlags AS DWORD = DOCHOSTUIFLAG_NO3DBORDER OR DOCHOSTUIFLAG_THEME OR DOCHOSTUIFLAG_DPI_AWARE)
   DECLARE DESTRUCTOR
   DECLARE FUNCTION hWindow () AS HWND
   DECLARE FUNCTION BrowserPtr () AS Afx_IWebBrowser2 PTR
   DECLARE PROPERTY DocHostFlags (BYVAL dwDocHostFlags AS DWORD)
   DECLARE PROPERTY DocHostFlags () AS DWORD
   DECLARE PROPERTY DocHostDoubleClickFlags (BYVAL dwDocHostDoubleClickFlags AS DWORD)
   DECLARE PROPERTY DocHostDoubleClickFlags () AS DWORD
   DECLARE PROPERTY AllowContextMenu (BYVAL bAllowContextMenu AS VARIANT_BOOL)
   DECLARE PROPERTY AllowContextMenu () AS VARIANT_BOOL
   DECLARE PROPERTY AllowShowUI (BYVAL bAllowShowUI AS VARIANT_BOOL)
   DECLARE PROPERTY AllowShowUI () AS VARIANT_BOOL
   DECLARE PROPERTY AllowWindowlessActivation (BYVAL bAllowWindowless AS VARIANT_BOOL)
   DECLARE PROPERTY AllowWindowlessActivation () AS VARIANT_BOOL
   DECLARE PROPERTY BackColor (BYVAL clrBackground AS OLE_COLOR)
   DECLARE PROPERTY BackColor () AS OLE_COLOR
   DECLARE PROPERTY ForeColor (BYVAL clrBackground AS OLE_COLOR)
   DECLARE PROPERTY ForeColor () AS OLE_COLOR
   DECLARE PROPERTY LocaleID (BYVAL lcidLocaleID AS LCID)
   DECLARE PROPERTY LocaleID () AS LCID
   DECLARE PROPERTY DisplayAsDefault (BYVAL bDisplayAsDefault AS VARIANT_BOOL)
   DECLARE PROPERTY DisplayAsDefault () AS VARIANT_BOOL
   DECLARE PROPERTY MessageReflect (BYVAL bMessageReflect AS VARIANT_BOOL)
   DECLARE PROPERTY MessageReflect () AS VARIANT_BOOL
   DECLARE PROPERTY Font (BYVAL pFont AS IFontDisp PTR)
   DECLARE PROPERTY Font () AS IFontDisp PTR
   DECLARE FUNCTION Navigate (BYVAL pwszUrl AS WSTRING PTR, BYVAL Flags AS VARIANT PTR = NULL, BYVAL TargetFrameName AS VARIANT PTR = NULL, BYVAL PostData AS VARIANT PTR = NULL, BYVAL Headers AS VARIANT PTR = NULL) AS HRESULT
   DECLARE FUNCTION Navigate (BYVAL pvUrl AS VARIANT PTR, BYVAL Flags AS VARIANT PTR = NULL, BYVAL TargetFrameName AS VARIANT PTR = NULL, BYVAL PostData AS VARIANT PTR = NULL, BYVAL Headers AS VARIANT PTR = NULL) AS HRESULT
   DECLARE FUNCTION WaitForPageLoad (BYVAL nTimeout AS DOUBLE = 10) AS READYSTATE
   DECLARE FUNCTION SetFocus () AS HRESULT
   DECLARE PROPERTY Silent (BYVAL pbSilent AS VARIANT_BOOL)
   DECLARE PROPERTY Silent () AS VARIANT_BOOL
   DECLARE PROPERTY Resizable (BYVAL bResize AS VARIANT_BOOL)
   DECLARE PROPERTY Resizable () AS VARIANT_BOOL
   DECLARE FUNCTION GoBack () AS HRESULT
   DECLARE FUNCTION GoForward () AS HRESULT
   DECLARE FUNCTION GoHome () AS HRESULT
   DECLARE FUNCTION GoSearch () AS HRESULT
   DECLARE FUNCTION Refresh () AS HRESULT
   DECLARE FUNCTION Refresh2 (BYVAL nLevel AS RefreshConstants) AS HRESULT
   DECLARE FUNCTION Stop () AS HRESULT
   DECLARE FUNCTION PrintPreview () AS HRESULT
   DECLARE FUNCTION PrintPage (BYVAL bPromptUser AS BOOLEAN = TRUE) AS HRESULT
   DECLARE FUNCTION PageSetUp () AS HRESULT
   DECLARE FUNCTION PageProperties () AS HRESULT
   DECLARE FUNCTION SaveAs () AS HRESULT
   DECLARE FUNCTION Find () AS HRESULT
   DECLARE FUNCTION ShowSource () AS HRESULT
   DECLARE FUNCTION InternetOptions () AS HRESULT
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
   DECLARE FUNCTION WriteHtml (BYREF cwsHtml AS CWSTR, BYVAL cr AS BOOLEAN = FALSE) AS HRESULT
   DECLARE FUNCTION GetBodyInnerText () AS CWSTR
   DECLARE FUNCTION GetBodyInnerHtml () AS CWSTR
   DECLARE FUNCTION GetActiveElementId () AS CWSTR
   DECLARE FUNCTION GetElementValueById (BYREF cwsId AS CWSTR) AS VARIANT
   DECLARE FUNCTION SetElementValueById (BYREF cwsId AS CWSTR, BYVAL vValue AS VARIANT) AS HRESULT
   DECLARE FUNCTION GetElementInnerHtmlById (BYREF cwsId AS CWSTR) AS CWSTR
   DECLARE FUNCTION SetElementInnerHtmlById (BYREF cwsId AS CWSTR, BYREF cwsHtml AS CWSTR) AS HRESULT
   DECLARE FUNCTION SetElementFocusById (BYREF cwsId AS CWSTR) AS HRESULT
   DECLARE FUNCTION Advise () AS HRESULT
   DECLARE FUNCTION Unadvise () AS HRESULT

END TYPE
' ========================================================================================

' ========================================================================================
' CWebCtx class constructors
' ========================================================================================
PRIVATE CONSTRUCTOR CWebCtx (BYREF pWindow AS CWindow, BYVAL cID AS INTEGER, BYREF wszURL AS WSTRING = "about:blank", _
BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
BYVAL dwStyle AS DWORD = WS_VISIBLE OR WS_CHILD OR WS_TABSTOP, BYVAL dwExStyle AS DWORD = 0, _
BYVAL bSilent AS VARIANT_BOOL = VARIANT_TRUE, BYVAL dwDocHostFlags AS DWORD = DOCHOSTUIFLAG_NO3DBORDER OR DOCHOSTUIFLAG_THEME OR DOCHOSTUIFLAG_DPI_AWARE)
   CATL_DP("CWebCtx Constructor CWindow")

   ' // Create the control
   m_hWb = pWindow.AddControl(m_wszAtlClassName, pWindow.hWindow, cID, "Shell Explorer", x, y, nWidth, nHeight, dwStyle, dwExStyle)
   ' // Store a pointer to the class as a property
   SetPropW(m_hWb, "CWEBCTXPTR", CAST(HANDLE, @this))
   ' // Get an addrefed pointer to the IWebBrowser2 interface
   m_pWb = CAST(Afx_IWebBrowser2 PTR, CAST(ULONG_PTR, this.GetDispatch(m_hWb)))
   ' // Make navigation silent (disables annoying popups of the script debugger)
   this.Silent = bSilent
   ' // Make the web page themed and DPI aware
   this.DocHostFlags = dwDocHostFlags
   ' // Navigate to the specified URL
   this.Navigate(wszURL)

END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CWebCtx (BYVAL pWindow AS CWindow PTR, BYVAL cID AS INTEGER, BYREF wszURL AS WSTRING = "about:blank", _
BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
BYVAL dwStyle AS DWORD = WS_VISIBLE OR WS_CHILD OR WS_TABSTOP, BYVAL dwExStyle AS DWORD = 0, _
BYVAL bSilent AS VARIANT_BOOL = VARIANT_TRUE, BYVAL dwDocHostFlags AS DWORD = DOCHOSTUIFLAG_NO3DBORDER OR DOCHOSTUIFLAG_THEME OR DOCHOSTUIFLAG_DPI_AWARE)
   CATL_DP("CWebCtx Constructor CWindow Ptr")

   ' // Create the control
   m_hWb = pWindow->AddControl(m_wszAtlClassName, pWindow->hWindow, cID, "Shell Explorer", x, y, nWidth, nHeight, dwStyle, dwExStyle)
   ' // Store a pointer to the class as a property
   SetPropW(m_hWb, "CWEBCTXPTR", CAST(HANDLE, @this))
   ' // Get an addrefed pointer to the IWebBrowser2 interface
   m_pWb = CAST(Afx_IWebBrowser2 PTR, CAST(ULONG_PTR, this.GetDispatch(m_hWb)))
   ' // Make navigation silent (disables annoying popups of the script debugger)
   this.Silent = bSilent
   ' // Make the web page themed and DPI aware
   this.DocHostFlags = dwDocHostFlags
   ' // Navigate to the specified URL
   this.Navigate(wszURL)

END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CWebCtx class destructor
' ========================================================================================
PRIVATE DESTRUCTOR CWebCtx
   ' // Clear the web page
   this.Navigate("about:blank")
   ' // Stop navigation
   this.Stop
   ' // Remove the window property
   RemovePropW(m_hWb, "CWEBCTXPTR")
   ' // Release the IWebBrowser2 inteface
   DIM r AS DWORD
   IF m_pWb THEN r = m_pWb->Release
   CATL_DP("CWebCtx Destructor - ref = " & WSTR(r))
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the handle of the control.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.hWindow () AS HWND
   CATL_DP("CWebCtx hWindow")
   RETURN m_hWb
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a direct pointer to the IWebBrowser2 interface.
' Note: Don't release the returned pointer.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.BrowserPtr () AS Afx_IWebBrowser2 PTR
   CATL_DP("CWebCtx BrowserPtr")
   RETURN m_pWb
END FUNCTION
' ========================================================================================

' ========================================================================================
' Specifies the user interface capabilities of the host object.
' ========================================================================================
PRIVATE PROPERTY CWebCtx.DocHostFlags (BYVAL dwDocHostFlags AS DWORD)
   CATL_DP("CWebCtx put_DocHostFlags")
   Base.DocHostFlags(m_hWb) = dwDocHostFlags
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CWebCtx.DocHostFlags () AS DWORD
'   CATL_DP("CWebCtx get_DocHostFlags")
   RETURN Base.DocHostFlags(m_hWb)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Specifies the operation that should take place in response to a double-click.
' ========================================================================================
PRIVATE PROPERTY CWebCtx.DocHostDoubleClickFlags (BYVAL dwDocHostDoubleClickFlags AS DWORD)
'   CATL_DP("CWebCtx put_DocHostDoubleClickFlags")
   Base.DocHostDoubleClickFlags(m_hWb) = dwDocHostDoubleClickFlags
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CWebCtx.DocHostDoubleClickFlags () AS DWORD
'   CATL_DP("CWebCtx get_DocHostDoubleClickFlags")
   RETURN Base.DocHostDoubleClickFlags(m_hWb)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Specifies whether the hosted control is allowed to display its own context menu.
' - bAllowContextMenu: VARIANT_TRUE (context menu allowed), VARIANT_FALSE (context menu disabled).
' The ATL host object implementation uses VARIANT_TRUE as the default value of this property.
' ========================================================================================
PRIVATE PROPERTY CWebCtx.AllowContextMenu (BYVAL bAllowContextMenu AS VARIANT_BOOL)
'   CATL_DP("CWebCtx put_AllowContextMenu")
   Base.AllowContextMenu(m_hWb) = bAllowContextMenu
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CWebCtx.AllowContextMenu () AS VARIANT_BOOL
'   CATL_DP("CWebCtx get_AllowContextMenu")
   RETURN Base.AllowContextMenu(m_hWb)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Specifies whether the hosted control is allowed to display its own user interface.
' - bAllowShowUI: VARIANT_TRUE or VARIANT_FALSE.
' The ATL host object implementation uses VARIANT_TRUE as the default value of this property.
' ========================================================================================
PRIVATE PROPERTY CWebCtx.AllowShowUI (BYVAL bAllowShowUI AS VARIANT_BOOL)
'   CATL_DP("CWebCtx put_AllowShowUI")
   Base.AllowShowUI(m_hWb) = bAllowShowUI
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CWebCtx.AllowShowUI () AS VARIANT_BOOL
'   CATL_DP("CWebCtx get_AllowShowUI")
   RETURN Base.AllowShowUI(m_hWb)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Specifies whether the container will allow windowless activation.
' - bAllowWindowless: VARIANT_TRUE or VARIANT_FALSE.
' The ATL host object implementation uses VARIANT_TRUE as the default value of this property.
' ========================================================================================
PRIVATE PROPERTY CWebCtx.AllowWindowlessActivation (BYVAL bAllowWindowless AS VARIANT_BOOL)
'   CATL_DP("CWebCtx put_AllowWindowlessActivation")
   Base.AllowWindowlessActivation(m_hWb) = bAllowWindowless
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CWebCtx.AllowWindowlessActivation () AS VARIANT_BOOL
'   CATL_DP("CWebCtx get_AllowWindowlessActivation")
   RETURN Base.AllowWindowlessActivation(m_hWb)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Specifies the ambient background color of the container.
' - clrBackground: The new value of this property.
' The ATL host object implementation uses COLOR_BTNFACE or COLOR_WINDOW as the default value
' of this property (depending on whether the parent of the host window is a dialog or not).
' ========================================================================================
PRIVATE PROPERTY CWebCtx.BackColor (BYVAL clrBackground AS OLE_COLOR)
'   CATL_DP("CWebCtx put_BackColor")
   Base.BackColor(m_hWb) = clrBackground
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CWebCtx.BackColor () AS OLE_COLOR
'   CATL_DP("CWebCtx get_BackColor")
   RETURN Base.BackColor(m_hWb)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Specifies the ambient foreground color of the container.
' - clrForeground: The new value of this property.
' The ATL host object implementation uses the system window text color as the default
' value of this property.
' ========================================================================================
PRIVATE PROPERTY CWebCtx.ForeColor (BYVAL clrForeground AS OLE_COLOR)
'   CATL_DP("CWebCtx put_ForeColor")
   Base.ForeColor(m_hWb) = clrForeground
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CWebCtx.ForeColor () AS OLE_COLOR
'   CATL_DP("CWebCtx get_ForeColor")
   RETURN Base.ForeColor(m_hWb)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Specifies the ambient locale ID of the container.
' - lcidLocaleID: The new value of this property.
' The ATL host object implementation uses the user's default locale as the default value
' of this property.
' ========================================================================================
PRIVATE PROPERTY CWebCtx.LocaleID (BYVAL lcidLocaleID AS LCID)
'   CATL_DP("CWebCtx put_LocaleID")
   Base.LocaleID(m_hWb) = lcidLocaleID
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CWebCtx.LocaleID () AS LCID
'   CATL_DP("CWebCtx get_LocaleID")
   RETURN Base.LocaleID(m_hWb)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Specifies the ambient locale ID of the container.
' - lcidLocaleID: The new value of this property.
' The ATL host object implementation uses the user's default locale as the default value
' of this property.
' ========================================================================================
PRIVATE PROPERTY CWebCtx.DisplayAsDefault (BYVAL bDisplayAsDefault AS VARIANT_BOOL)
'   CATL_DP("CWebCtx put_DisplayAsDefault")
   Base.DisplayAsDefault(m_hWb) = bDisplayAsDefault
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CWebCtx.DisplayAsDefault () AS VARIANT_BOOL
'   CATL_DP("CWebCtx get_DisplayAsDefault")
   RETURN Base.DisplayAsDefault(m_hWb)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Specifies whether the container will reflect messages to the hosted control.
' - bMessageReflect: VARIANT_TRUE or VARIANT_FALSE.
' The ATL host object implementation uses VARIANT_TRUE as the default value of this property.
' ========================================================================================
PRIVATE PROPERTY CWebCtx.MessageReflect (BYVAL bMessageReflect AS VARIANT_BOOL)
'   CATL_DP("CWebCtx put_MessageReflect")
   Base.MessageReflect(m_hWb) = bMessageReflect
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CWebCtx.MessageReflect () AS VARIANT_BOOL
'   CATL_DP("CWebCtx get_MessageReflect")
   RETURN Base.MessageReflect(m_hWb)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Specifies the ambient font of the container.
' - pFont: Pointer to the IFontDisp object to set.
' The ATL host object implementation uses the default GUI font or the system font as the
' default value of this property.
' ========================================================================================
PRIVATE PROPERTY CWebCtx.Font (BYVAL pFont AS IFontDisp PTR)
'   CATL_DP("CWebCtx put_Font")
   Base.Font(m_hWb) = pFont
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CWebCtx.Font () AS IFontDisp PTR
'   CATL_DP("CWebCtx get_Font")
   RETURN Base.Font(m_hWb)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Navigates to the specified URL.
' Note: To clear a web page, simply navigate to "about:blank".
' ========================================================================================
PRIVATE FUNCTION CWebCtx.Navigate (BYVAL pwszUrl AS WSTRING PTR, BYVAL Flags AS VARIANT PTR = NULL, BYVAL TargetFrameName AS VARIANT PTR = NULL, BYVAL PostData AS VARIANT PTR = NULL, BYVAL Headers AS VARIANT PTR = NULL) AS HRESULT
   CATL_DP("CWebCtx Navigate pwszUrl")
   IF m_pWB = NULL THEN RETURN E_POINTER
   DIM vUrl AS VARIANT : vUrl.vt = VT_BSTR : vUrl.bstrVal = SysAllocString(pwszUrl)
   DIM hr AS HRESULT = m_pWb->Navigate2(@vUrl, Flags, TargetFrameName, PostData, Headers)
   VariantClear @vurl
   RETURN hr
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CWebCtx.Navigate (BYVAL pvUrl AS VARIANT PTR, BYVAL Flags AS VARIANT PTR = NULL, BYVAL TargetFrameName AS VARIANT PTR = NULL, BYVAL PostData AS VARIANT PTR = NULL, BYVAL Headers AS VARIANT PTR = NULL) AS HRESULT
   CATL_DP("CWebCtx Navigate Variant Ptr")
   IF m_pWB = NULL THEN RETURN E_POINTER
   DIM hr AS HRESULT = m_pWb->Navigate2(pvUrl, Flags, TargetFrameName, PostData, Headers)
   RETURN hr
END FUNCTION
' ========================================================================================

' ========================================================================================
' Waits until the page had been fully downloaded or te timeout has expired.
' - dbTimeout Timeout in seconds.
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
   CATL_DP("CWebCtx WaitForPageLoad")
   DIM lReadyState AS tagREADYSTATE
   DIM starttime AS DOUBLE = timer
   DO
      ' // Processes pending messages
      ' // Needed to allow the page to load
      AfxPumpMessages
      ' // Retrieves the ready state
      m_pWb->get_ReadyState(@lReadyState)
      IF lReadyState = READYSTATE_COMPLETE THEN EXIT DO
      IF timer - starttime > dbTimeout THEN EXIT DO
      ' // Needed to update timer
      SLEEP 1, 1
   LOOP
   RETURN lReadyState
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the focus in the hosted document.
' Return value: S_OK if successful, or an error value otherwise.
' Note: get_Document will return S_FALSE if the control has not finished loading the new
' document and all its contents. Therefore, trying to set the focus too early will fail.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.SetFocus () AS HRESULT
   CATL_DP("CWebCtx SetFocus")
   IF m_pWB = NULL THEN RETURN E_POINTER
   DIM pIHTMLDocument2 AS IHTMLDocument2 PTR
   DIM hr AS HRESULT =  m_pWb->get_Document(@CAST(ANY PTR, pIHTMLDocument2))
   IF hr <> S_OK THEN RETURN hr
   DIM pWindow2 AS IHTMLWindow2 PTR
   hr = pIHTMLDocument2->lpvtbl->get_parentWindow(pIHTMLDocument2, @pWindow2)
   pIHTMLDocument2->lpvtbl->Release(pIHTMLDocument2)
   IF hr <> S_OK THEN RETURN hr
   hr = pWindow2->lpvtbl->focus(pWindow2)
   pWindow2->lpvtbl->Release(pWindow2)
   RETURN hr
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets or gets a value that indicates whether the object can display dialog boxes.
' VARIANT_FALSE Default. Dialog boxes and messages can be displayed.
' VARIANT_TRUE  Dialog boxes are not displayed.
' Note: Critical errors and security alerts are not suppressed.
' ========================================================================================
PRIVATE PROPERTY CWebCtx.Silent (BYVAL bSilent AS VARIANT_BOOL)
   CATL_DP("CWebCtx put_Silent")
   IF m_pWb THEN m_pWb->put_Silent(bSilent)
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CWebCtx.Silent () AS VARIANT_BOOL
   CATL_DP("CWebCtx get_Silent")
   DIM bSilent AS VARIANT_BOOL
   IF m_pWb THEN m_pWb->get_Silent(@bSilent)
   RETURN bSilent
END PROPERTY
' ========================================================================================

' ========================================================================================
' Sets or gets a value that indicates whether the object can be resized.
' ========================================================================================
PRIVATE PROPERTY CWebCtx.Resizable (BYVAL bResize AS VARIANT_BOOL)
   CATL_DP("CWebCtx put_Resizable")
   IF m_pWb THEN m_pWb->put_Resizable(bResize)
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CWebCtx.Resizable () AS VARIANT_BOOL
   CATL_DP("CWebCtx get_Resizable")
   DIM bResize AS VARIANT_BOOL
   IF m_pWb THEN m_pWb->get_Resizable(@bResize)
   RETURN bResize
END PROPERTY
' ========================================================================================

' ========================================================================================
' Navigates backward one item in the history list.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.GoBack () AS HRESULT
   IF m_pWb THEN RETURN m_pWb->GoBack
END FUNCTION
' ========================================================================================
' ========================================================================================
' Navigates forward one item in the history list.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.GoForward () AS HRESULT
   IF m_pWb THEN RETURN m_pWb->GoForward
END FUNCTION
' ========================================================================================
' ========================================================================================
' Navigates to the current home or start page.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.GoHome () AS HRESULT
   IF m_pWb THEN RETURN m_pWb->GoHome
END FUNCTION
' ========================================================================================
' ========================================================================================
' Navigates to the default search page of the current user.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.GoSearch () AS HRESULT
   IF m_pWb THEN RETURN m_pWb->GoSearch
END FUNCTION
' ========================================================================================
' ========================================================================================
' Reloads the file that is currently displayed in the object.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.Refresh () AS HRESULT
   IF m_pWb THEN RETURN m_pWb->Refresh
END FUNCTION
' ========================================================================================
' ========================================================================================
' Reloads the file that is currently displayed with the specified refresh level.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.Refresh2 (BYVAL nLevel AS RefreshConstants) AS HRESULT
   DIM vLevel AS VARIANT
   vLevel.vt = VT_I4 : vLevel.lVal = nLevel
   IF m_pWb THEN RETURN m_pWb->Refresh2(@vLevel)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Activates the print preview dialog.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.PrintPreview () AS HRESULT
   IF m_pWb THEN RETURN  m_pWb->ExecWB(OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_PROMPTUSER, NULL, NULL)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Activates de print dialog.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.PrintPage (BYVAL bPromptUser AS BOOLEAN = TRUE) AS HRESULT
   IF m_pWb = NULL THEN RETURN E_POINTER
   IF bPromptUser THEN
      m_pWb->ExecWB(OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER, NULL, NULL)
   ELSE
      m_pWb->ExecWB(OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER, NULL, NULL)
   END IF
END FUNCTION
' ========================================================================================
' ========================================================================================
' Activates the page setup dialog.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.PageSetUp () AS HRESULT
   IF m_pWb THEN RETURN  m_pWb->ExecWB(OLECMDID_PAGESETUP, OLECMDEXECOPT_PROMPTUSER, NULL, NULL)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Activates the properties dialog.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.PageProperties () AS HRESULT
   IF m_pWb THEN RETURN  m_pWb->ExecWB(OLECMDID_PROPERTIES, OLECMDEXECOPT_PROMPTUSER, NULL, NULL)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Activates the save file dialog.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.SaveAs () AS HRESULT
   IF m_pWb THEN RETURN  m_pWb->ExecWB(OLECMDID_SAVEAS, OLECMDEXECOPT_PROMPTUSER, NULL, NULL)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Activates the find dialog.
' Warning: This code uses an undocumented command-group GUID that is subject to change in
' the future. Currently it works in all versions of Internet Explorer up to 11.
' See http://support.microsoft.com/?kbid=311288
' ========================================================================================
PRIVATE FUNCTION CWebCtx.Find () AS HRESULT
   IF m_pWb = NULL THEN RETURN E_POINTER
   DIM pIHTMLDocument2 AS IHTMLDocument2 PTR
   DIM hr AS HRESULT = m_pWb->get_Document(@CAST(ANY PTR, pIHTMLDocument2))
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
   IF m_pWb = NULL THEN RETURN E_POINTER
   DIM pIHTMLDocument2 AS IHTMLDocument2 PTR
   DIM hr AS HRESULT = m_pWb->get_Document(@CAST(ANY PTR, pIHTMLDocument2))
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
   IF m_pWb = NULL THEN RETURN E_POINTER
   DIM pIHTMLDocument2 AS IHTMLDocument2 PTR
   DIM hr AS HRESULT = m_pWb->get_Document(@CAST(ANY PTR, pIHTMLDocument2))
   IF hr <> S_OK OR pIHTMLDocument2 = NULL THEN RETURN hr
   DIM CGID_WebBrowser_ AS GUID = (&hED016940, &hBD5B, &h11CF, {&hBA, &h4E, &h00, &hC0, &h4F, &hD7, &h08, &h16})
   DIM pCmdTarget AS IOleCommandTarget PTR
   hr = pIHTMLDocument2->lpvtbl->QueryInterface(pIHTMLDocument2, @IID_IOleCommandTarget, @pCmdTarget)
   pIHTMLDocument2->lpvtbl->Release(pIHTMLDocument2)
   IF hr <> S_OK OR pIHTMLDocument2 = NULL THEN RETURN hr
   hr = pCmdTarget->lpvtbl->Exec(pCmdTarget, @CGID_WebBrowser_, 3, 0, NULL, NULL)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Cancels a pending navigation or download, and stops dynamic page elements, such as
' background sounds and animations.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.Stop () AS HRESULT
   IF m_pWb THEN RETURN m_pWb->Stop
END FUNCTION
' ========================================================================================
' ========================================================================================
' Gets the automation object of the active document, if any.
' Release the returned pointer when no longer needed.
' ========================================================================================
PRIVATE PROPERTY CWebCtx.Document () AS IHtmlDocument2 PTR
   DIM pdisp AS IDispatch PTR
   IF m_pWb THEN m_pWb->get_Document(CAST(ANY PTR, @pdisp))
   PROPERTY = CAST(IHtmlDocument2 PTR, pdisp)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Retrieves the path or title of the resource that is currently displayed.
' ========================================================================================
PRIVATE PROPERTY CWebCtx.LocationName () AS CWSTR
   DIM bstrName AS AFX_BSTR
   IF m_pWb THEN m_pWb->get_LocationName(@bstrName)
   PROPERTY = *bstrName
   SysFreestring bstrName
END PROPERTY
' ========================================================================================
' ========================================================================================
' Retrieves the URL of the resource that is currently displayed.
' ========================================================================================
PRIVATE PROPERTY CWebCtx.LocationURL () AS CWSTR
   DIM bstrName AS AFX_BSTR
   IF m_pWb THEN m_pWb->get_LocationURL(@bstrName)
   PROPERTY = *bstrName
   SysFreestring bstrName
END PROPERTY
' ========================================================================================
' ========================================================================================
' TRUE if the object is engaged in a navigation or downloading operation; FALSE, otherwise.
' ========================================================================================
PRIVATE PROPERTY CWebCtx.Busy () AS BOOLEAN
   DIM bBusy AS VARIANT_BOOL
   IF m_pWb THEN m_pWb->get_Busy(@bBusy)
   PROPERTY = bBusy
END PROPERTY
' ========================================================================================
' ========================================================================================
' Queries the object for the status of commands using the IOleCommandTarget::QueryStatus method.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.QueryStatusWB (BYVAL cmdID AS OLECMDID) AS OLECMDF
   DIM cmdf AS OLECMDF
   IF m_pWb THEN m_pWb->QueryStatusWB(cmdID, @cmdf)
   RETURN cmdf
END FUNCTION
' ========================================================================================
' ========================================================================================
' Executes a command and returns the status of the command execution using the IOleCommandTarget interface.
' ========================================================================================
PRIVATE FUNCTION CWebCtx.ExecWB (BYVAL cmdID AS OLECMDID, BYVAL cmdexecopt AS OLECMDEXECOPT, BYVAL pvaIn AS VARIANT PTR = NULL, BYVAL pvaOut AS VARIANT PTR = NULL) AS HRESULT
   IF m_pWb THEN RETURN(m_pWb->ExecWB(cmdID, cmdexecopt, pvaIn, pvaOut))
END FUNCTION
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
   DIM nState AS tagREADYSTATE
   IF m_pWb THEN m_pWb->get_ReadyState(@nState)
   PROPERTY = nState
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets or gets a value that indicates whether the object is registered as a top-level browser window.
' bRegister: VARIANT_TRUE or VARIANT_FALSE
' ========================================================================================
PRIVATE PROPERTY CWebCtx.RegisterAsBrowser () AS BOOLEAN
   DIM bRegister AS VARIANT_BOOL
   IF m_pWb THEN m_pWb->get_RegisterAsBrowser(@bRegister)
   PROPERTY = bRegister
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CWebCtx.RegisterAsBrowser (BYVAL bRegister AS BOOLEAN)
   IF m_pWb THEN m_pWb->put_RegisterAsBrowser(bRegister)
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets or gets a value that indicates whether the object is registered as a drop target for navigation.
' bRegister: VARIANT_TRUE or VARIANT_FALSE (disables drag and drop)
' ========================================================================================
PRIVATE PROPERTY CWebCtx.RegisterAsDropTarget () AS BOOLEAN
   DIM bRegister AS VARIANT_BOOL
   IF m_pWb THEN m_pWb->get_RegisterAsDropTarget(@bRegister)
   PROPERTY = bRegister
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CWebCtx.RegisterAsDropTarget (BYVAL bRegister AS BOOLEAN)
   IF m_pWb THEN m_pWb->put_RegisterAsDropTarget(bRegister)
END PROPERTY
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
   IF m_pWb = NULL THEN RETURN E_POINTER
   DIM pIHTMLDocument2 AS IHTMLDocument2 PTR
   DIM hr AS HRESULT = m_pWb->get_Document(@CAST(ANY PTR, pIHTMLDocument2))
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
   IF m_pWb = NULL THEN RETURN ""
   DIM pIHTMLDocument2 AS IHTMLDocument2 PTR
   DIM hr AS HRESULT = m_pWb->get_Document(@CAST(ANY PTR, pIHTMLDocument2))
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
   IF m_pWb = NULL THEN RETURN ""
   DIM pIHTMLDocument2 AS IHTMLDocument2 PTR
   DIM hr AS HRESULT = m_pWb->get_Document(@CAST(ANY PTR, pIHTMLDocument2))
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
   IF m_pWb = NULL THEN RETURN ""
   DIM pIHTMLDocument2 AS IHTMLDocument2 PTR
   DIM hr AS HRESULT = m_pWb->get_Document(@CAST(ANY PTR, pIHTMLDocument2))
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
   IF m_pWb = NULL THEN EXIT FUNCTION
   ' // Get a reference to the IHTMLDocument2 interface
   DIM pIHTMLDocument2 AS IHTMLDocument2 PTR
   DIM hr AS HRESULT = m_pWb->get_Document(@CAST(ANY PTR, pIHTMLDocument2))
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
   IF m_pWb = NULL THEN RETURN E_POINTER
   ' // Get a reference to the IHTMLDocument2 interface
   DIM pIHTMLDocument2 AS IHTMLDocument2 PTR
   DIM hr AS HRESULT = m_pWb->get_Document(@CAST(ANY PTR, pIHTMLDocument2))
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
   IF m_pWb = NULL THEN RETURN ""
   ' // Get a reference to the IHTMLDocument2 interface
   DIM pIHTMLDocument2 AS IHTMLDocument2 PTR
   DIM hr AS HRESULT = m_pWb->get_Document(@CAST(ANY PTR, pIHTMLDocument2))
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
   IF m_pWb = NULL THEN RETURN E_POINTER
   ' // Get a reference to the IHTMLDocument2 interface
   DIM pIHTMLDocument2 AS IHTMLDocument2 PTR
   DIM hr AS HRESULT = m_pWb->get_Document(@CAST(ANY PTR, pIHTMLDocument2))
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
   IF m_pWb = NULL THEN RETURN E_POINTER
   ' // Get a reference to the IHTMLDocument2 interface
   DIM pIHTMLDocument2 AS IHTMLDocument2 PTR
   DIM hr AS HRESULT = m_pWb->get_Document(@CAST(ANY PTR, pIHTMLDocument2))
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
' Returns the ATL container window handle given the handle of the form, or any control in the
' form, and its control identifier.
' Parameters:
' - hwnd = Reference window handle.
' - Control identifier, e.g. IDC_WEBBROWSER.
' Returns the ATL container window handle or NULL.
' ========================================================================================
PRIVATE FUNCTION AtlControlWindow (BYVAL hwnd AS HWND, BYVAL cID AS WORD) AS HWND
   DIM wszClassName AS WSTRING * 260
   DIM hwndChild AS .HWND = AfxGetControlHandle(hwnd, cID)
   GetClassNameW hwndChild, wszClassName, SIZEOF(wszClassName)
   IF wszClassName <> "AtlAxWin" THEN RETURN NULL
   RETURN hwndChild
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a direct IUnknown pointer to the hosted control given the handle of the
' control that hosts it.
' - hwnd = The ATL control handle
' Don't release the returned pointer because it is a direct pointer.
' ========================================================================================
PRIVATE FUNCTION AtlControlUnknown (BYVAL hwnd AS HWND) AS Afx_IUnknown PTR
   DIM hAtlLib AS HMODULE
   hAtlLib = LoadLibraryW("Atl.dll")
   IF hAtlLib = NULL THEN RETURN NULL
   DIM pProc AS FUNCTION (BYVAL hwnd AS .HWND, BYREF punk AS IUnknown PTR) AS HRESULT
   pProc = CAST(ANY PTR, GetProcAddress(hAtlLib, "AtlAxGetControl"))
   FreeLibrary hAtlLib
   IF pProc = NULL THEN RETURN NULL
   DIM punk AS IUnknown PTR
   DIM hr AS HRESULT = pProc(hwnd, punk)
   RETURN CAST(Afx_IUnknown PTR, CAST(ULONG_PTR, punk))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns a direct IDispatch pointer to the hosted control given the handle of the
' control that hosts it.
' - hwnd = The ATL control handle
' Don't release the returned pointer because it is a direct pointer.
' ========================================================================================
PRIVATE FUNCTION AtlControlDispatch (BYVAL hwnd AS HWND) AS Afx_IDispatch PTR
   DIM punk AS Afx_IUnknown PTR = AtlControlUnknown(hwnd)
   IF punk = NULL THEN RETURN NULL
   ' // Query for the IDispatch interface
   DIM pdisp AS Afx_IDispatch PTR
   punk->QueryInterface(@IID_IDispatch, @pdisp)
   ' // Decrease the reference count to make it a direct pointer
   pDisp->Release
   RETURN pdisp
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns a reference counted pointer to the WebBrowser control given the handle of the
' control that hosts it.
' - hwnd = The ATL control handle
' Release the returned pointer when you no longer need it.
' ========================================================================================
PRIVATE FUNCTION AtlControlObject (BYVAL hwnd AS HWND) AS Afx_IDispatch PTR
   DIM punk AS Afx_IUnknown PTR = AtlControlUnknown(hwnd)
   IF punk = NULL THEN RETURN NULL
   ' // Query for the IDispatch interface
   DIM pdisp AS Afx_IDispatch PTR
   punk->QueryInterface(@IID_IDispatch, @pdisp)
   RETURN pdisp
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a direct pointer to the WebBrowser control given the handle of the control
' that hosts it.
' - hwnd = The ATL control handle
' Don't release the returned pointer because it is a direct pointer.
' ========================================================================================
PRIVATE FUNCTION AtlBrowserDispatch OVERLOAD (BYVAL hwnd AS HWND) AS Afx_IWebBrowser2 PTR
   RETURN CAST(Afx_IWebBrowser2 PTR, CAST(ULONG_PTR, AtlControlDispatch(hwnd)))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns a direct pointer to the WebBrowser control given the handle of the form, or any
' control in the form, and its control identifier.
' Parameters:
' - hwnd = Reference window handle.
' - Control identifier, e.g. IDC_WEBBROWSER.
' Don't release the returned pointer because it is a direct pointer.
' ========================================================================================
PRIVATE FUNCTION AtlBrowserDispatch OVERLOAD (BYVAL hwnd AS HWND, BYVAL cID AS WORD) AS Afx_IWebBrowser2 PTR
   DIM hWb AS .HWND = AtlControlWindow(hwnd, cID)
   IF hWb = NULL THEN RETURN NULL
   RETURN AtlBrowserDispatch(hWb)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns a reference counted pointer to the WebBrowser control given the handle of the
' control that hosts it.
' - hwnd = The ATL control handle
' You must release the returned pointer when you no longer need it.
' ========================================================================================
PRIVATE FUNCTION AtlBrowserObject OVERLOAD (BYVAL hwnd AS HWND) AS Afx_IWebBrowser2 PTR
   RETURN CAST(Afx_IWebBrowser2 PTR, CAST(ULONG_PTR, AtlControlObject(hwnd)))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns a reference counted pointer to the WebBrowser control given the handle of the
' form, or any control in the form, and its control identifier.
' Parameters:
' - hwnd = Reference window handle.
' - Control identifier, e.g. IDC_WEBBROWSER.
' Don't release the returned pointer because it is a direct pointer.
' ========================================================================================
PRIVATE FUNCTION AtlBrowserObject OVERLOAD (BYVAL hwnd AS HWND, BYVAL cID AS WORD) AS Afx_IWebBrowser2 PTR
   DIM hWb AS .HWND = AtlControlWindow(hwnd, cID)
   IF hWb = NULL THEN RETURN NULL
   RETURN AtlBrowserObject(hWb)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a direct pointer to the CWebCtx class given the handle of the window of the ATL
' control created with it.
' ========================================================================================
PRIVATE FUNCTION AtlCWebCtx OVERLOAD (BYVAL hwnd AS HWND) AS CWebCtx PTR
   IF hwnd = NULL THEN RETURN NULL
   RETURN CAST(CWebCtx PTR, GetPropW(hwnd, "CWEBCTXPTR"))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns a direct pointer to the CWebCtx class given the handle of the form, or any
' control in the form, and its control identifier.
' ========================================================================================
PRIVATE FUNCTION AtlCWebCtx OVERLOAD (BYVAL hwnd AS HWND, BYVAL cID AS WORD) AS CWebCtx PTR
   DIM hWb AS .HWND = AtlControlWindow(hwnd, cID)
   IF hWb = NULL THEN RETURN NULL
   RETURN CAST(CWebCtx PTR, GetPropW(hWb, "CWEBCTXPTR"))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a standard IFontDisp object.
' It returns a pointer to a new IFontDisp object on success, or NULL on failure.
' Release the returned interface pointer when the new object is no longer needed. 
' ========================================================================================
PRIVATE FUNCTION AtlCreateFontDisp ( _
   BYREF wszFontName AS WSTRING, _            ' __in  Font name
   BYVAL cySize AS LONGLONG, _                ' __in  Point size
   BYVAL fWeight AS SHORT, _                  ' __in  Weight (%FW_NORMAL, etc.)
   BYVAL fCharset AS SHORT, _                 ' __in  Character set
   BYVAL fItalic AS WINBOOL, _                ' __in  Italic state
   BYVAL fUnderline AS WINBOOL, _             ' __in  Underline state
   BYVAL fStrikethrough AS WINBOOL _          ' __in  Strikethrough state
   ) AS IFontDisp PTR

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
' Forwards the message to the control. Active in-place objects must always be given the
' first chance at translating accelerator keystrokes.
' Usage example:
'   ' // Dispatch Windows messages
'   DIM uMsg AS MSG
'   WHILE (GetMessageW(@uMsg, NULL, 0, 0) <> FALSE)
'      IF AtlForwardMessage(GetFocus, @uMsg) = FALSE THEN
'         IF IsDialogMessageW(hWndMain, @uMsg) = 0 THEN
'            TranslateMessage(@uMsg)
'            DispatchMessageW(@uMsg)
'         END IF
'      END IF
'   WEND
' Note: WM_FORWARDMSG = &H37F ' (895)
' ========================================================================================
PRIVATE FUNCTION AtlForwardMessage (BYVAL hctl AS HWND, BYVAL pMsg AS tagMsg PTR) AS BOOLEAN
   IF (pMsg->message < WM_KEYFIRST OR pMsg->message > WM_KEYLAST) AND _
      (pMsg->message < WM_MOUSEFIRST OR pMsg->message > WM_MOUSELAST) THEN RETURN FALSE
   DIM hwnd AS HWND = hctl
   DIM wszClassName AS WSTRING * 260
   WHILE (GetWindowLongPtrW(hwnd, GWL_STYLE) AND WS_CHILD)
      IF (GetWindowLongPtrW(hwnd, GWL_EXSTYLE) AND WS_EX_MDICHILD) THEN EXIT WHILE
      GetClassNameW hwnd, wszClassName, SIZEOF(wszClassName)
      IF wszClassName = "AtlAxWin" THEN EXIT WHILE
      hwnd = GetParent(hwnd)
   WEND
   IF wszClassName <> "AtlAxWin" THEN RETURN FALSE
   RETURN SendMessageW(hwnd, &H37F, 0, CAST(LPARAM, pMsg))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Waits until the page had been fully downloaded or te timeout has expired.
' - dbTimeout Timeout in seconds.
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
PRIVATE FUNCTION AtlWaitForPageLoad (BYVAL hCtl AS HWND, BYVAL dbTimeout AS DOUBLE = 10) AS tagREADYSTATE
   DIM lReadyState AS tagREADYSTATE
   DIM starttime AS DOUBLE = timer
   DIM pWb AS Afx_IWebBrowser2 PTR = AtlBrowserDispatch(hCtl)
   IF pWb = NULL THEN RETURN 0
   DO
      ' // Processes pending messages
      ' // Needed to allow the page to load
      AfxPumpMessages
      ' // Retrieves the ready state
      pWb->get_ReadyState(@lReadyState)
      IF lReadyState = READYSTATE_COMPLETE THEN EXIT DO
      IF timer - starttime > dbTimeout THEN EXIT DO
      ' // Needed to update timer
      SLEEP 1, 1
   LOOP
   RETURN lReadyState
END FUNCTION
' ========================================================================================

END NAMESPACE
