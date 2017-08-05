' ########################################################################################
' Microsoft Windows
' File: AfxExDisp.inc
' Library name: SHDocVw
' Version: 1.1, Locale ID = 0
' Library GUID: LIBID_SHDocVw = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}"
' Path: C:\Windows\System32\ieframe.dll
' Contents: Microsoft Internet Controls
' Portions Copyright (c) Microsoft Corporation.
' Compiler: Free Basic 32 & 64 bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "win/exdisp.bi"
#include once "win/shtypes.bi"
#include once "win/mshtmhst.bi"

NAMESPACE Afx

' // The definition for BSTR in the FreeBASIC headers was inconveniently changed to WCHAR
#ifndef AFX_BSTR
   #define AFX_BSTR WSTRING PTR
#endif

' ========================================================================================
' ProgIDs (Program identifiers)
' ========================================================================================

' CLSID = {0002DF01-0000-0000-C000-000000000046}
CONST AFX_PROGID_InternetExplorer1 = "InternetExplorer.Application.1"
' CLSID = {EAB22AC3-30C1-11CF-A7EB-0000C05BAE0B}
CONST AFX_PROGID_WebBrowser_V11 = "Shell.Explorer.1"
' CLSID = {8856F961-340A-11D0-A96B-00C04FD705A2}
CONST AFX_PROGID_WebBrowser2 = "Shell.Explorer.2"
' CLSID = {64AB4BB7-111E-11D1-8F79-00C04FC2FBE1}
CONST AFX_PROGID_ShellUIHelper1 = "Shell.UIHelper.1"
' CLSID = {55136805-B2DE-11D1-B9F2-00A0C98BC547}
CONST AFX_PROGID_ShellNameSpace1 = "ShellNameSpace.ShellNameSpace.1"

' ========================================================================================
' Version Independent ProgIDs
' ========================================================================================

' CLSID = {0002DF01-0000-0000-C000-000000000046}
CONST AFX_PROGID_InternetExplorer = "InternetExplorer.Application"
' CLSID = {EAB22AC3-30C1-11CF-A7EB-0000C05BAE0B}
CONST AFX_PROGID_WebBrowser_V1 = "Shell.Explorer"
' CLSID = {55136805-B2DE-11D1-B9F2-00A0C98BC547}
CONST AFX_PROGID_ShellNameSpace = "ShellNameSpace.ShellNameSpace"

' ========================================================================================
' ClsIDs (Class identifiers)
' ========================================================================================

CONST AFX_CLSID_CScriptErrorList = "{EFD01300-160F-11D2-BB2E-00805FF7EFCA}"
CONST AFX_CLSID_InternetExplorer = "{0002DF01-0000-0000-C000-000000000046}"
CONST AFX_CLSID_InternetExplorerMedium = "{D5E8041D-920F-45E9-B8FB-B1DEB82C6E5E}"
CONST AFX_CLSID_ShellBrowserWindow = "{C08AFD90-F2A1-11D1-8455-00A0C91F3880}"
CONST AFX_CLSID_ShellNameSpace = "{55136805-B2DE-11D1-B9F2-00A0C98BC547}"
CONST AFX_CLSID_ShellUIHelper = "{64AB4BB7-111E-11D1-8F79-00C04FC2FBE1}"
CONST AFX_CLSID_ShellWindows = "{9BA05972-F6A8-11CF-A442-00A0C90A8F39}"
CONST AFX_CLSID_WebBrowser = "{8856F961-340A-11D0-A96B-00C04FD705A2}"
CONST AFX_CLSID_WebBrowser_V1 = "{EAB22AC3-30C1-11CF-A7EB-0000C05BAE0B}"

' ========================================================================================
' IIDs (Interface identifiers)
' ========================================================================================

CONST AFX_IID_DShellNameSpaceEvents = "{55136806-B2DE-11D1-B9F2-00A0C98BC547}"
CONST AFX_IID_DShellWindowsEvents = "{FE4106E0-399A-11D0-A48C-00A0C90A8F39}"
CONST AFX_IID_DWebBrowserEvents = "{EAB22AC2-30C1-11CF-A7EB-0000C05BAE0B}"
CONST AFX_IID_DWebBrowserEvents2 = "{34A715A0-6587-11D0-924A-0020AFC7AC4D}"
CONST AFX_IID_IScriptErrorList = "{F3470F24-15FD-11D2-BB2E-00805FF7EFCA}"
CONST AFX_IID_IShellFavoritesNameSpace = "{55136804-B2DE-11D1-B9F2-00A0C98BC547}"
CONST AFX_IID_IShellNameSpace = "{E572D3C9-37BE-4AE2-825D-D521763E3108}"
CONST AFX_IID_IShellUIHelper = "{729FE2F8-1EA8-11D1-8F85-00C04FC2FBE1}"
CONST AFX_IID_IShellUIHelper2 = "{A7FE6EDA-1932-4281-B881-87B31B8BC52C}"
CONST AFX_IID_IShellUIHelper3 = "{528DF2EC-D419-40BC-9B6D-DCDBF9C1B25D}"
CONST AFX_IID_IShellUIHelper4 = "{B36E6A53-8073-499E-824C-D776330A333E}"
CONST AFX_IID_IShellUIHelper5 = "{A2A08B09-103D-4D3F-B91C-EA455CA82EFA}"
CONST AFX_IID_IShellUIHelper6 = "{987A573E-46EE-4E89-96AB-DDF7F8FDC98C}"
CONST AFX_IID_IShellUIHelper7 = "{66DEBCF2-05B0-4F07-B49B-B96241A65DB2}"
CONST AFX_IID_IShellWindows = "{85CB6900-4D95-11CF-960C-0080C7F4EE85}"
CONST AFX_IID_IWebBrowser = "{EAB22AC1-30C1-11CF-A7EB-0000C05BAE0B}"
CONST AFX_IID_IWebBrowser2 = "{D30C1661-CDAF-11D0-8A3E-00C04FC9E26E}"
CONST AFX_IID_IWebBrowserApp = "{0002DF05-0000-0000-C000-000000000046}"

#ifndef __Afx_IUnknown_INTERFACE_DEFINED__
#define __Afx_IUnknown_INTERFACE_DEFINED__
TYPE Afx_IUnknown AS Afx_IUnknown_
TYPE Afx_IUnknown_ EXTENDS OBJECT
	DECLARE ABSTRACT FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
	DECLARE ABSTRACT FUNCTION AddRef () AS ULONG
	DECLARE ABSTRACT FUNCTION Release () AS ULONG
END TYPE
TYPE AFX_LPUNKNOWN AS Afx_IUnknown PTR
#endif

#ifndef __Afx_IDispatch_INTERFACE_DEFINED__
#define __Afx_IDispatch_INTERFACE_DEFINED__
TYPE Afx_IDispatch AS Afx_IDispatch_
TYPE Afx_IDispatch_  EXTENDS Afx_Iunknown
   DECLARE ABSTRACT FUNCTION GetTypeInfoCount (BYVAL pctinfo AS UINT PTR) as HRESULT
   DECLARE ABSTRACT FUNCTION GetTypeInfo (BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetIDsOfNames (BYVAL riid AS CONST IID CONST PTR, BYVAL rgszNames AS LPOLESTR PTR, BYVAL cNames AS UINT, BYVAL lcid AS LCID, BYVAL rgDispId AS DISPID PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL dispIdMember AS DISPID, BYVAL riid AS CONST IID CONST PTR, BYVAL lcid AS LCID, BYVAL wFlags AS WORD, BYVAL pDispParams AS DISPPARAMS PTR, BYVAL pVarResult AS VARIANT PTR, BYVAL pExcepInfo AS EXCEPINFO PTR, BYVAL puArgErr AS UINT PTR) AS HRESULT
END TYPE
TYPE AFX_LPDISPATCH AS Afx_IDispatch PTR
#endif

' ########################################################################################
' Interface name = IWebBrowser
' IID = "{EAB22AC1-30C1-11CF-A7EB-0000C05BAE0B}"
' Inherited interface = IDispatch
' Deprecated. Use IWebBrowser2.
' ########################################################################################

#define __Afx_IWebBrowser_INTERFACE_DEFINED__
TYPE Afx_IWebBrowser AS Afx_IWebBrowser_

TYPE Afx_IWebBrowser_ EXTENDS Afx_IDispatch
	DECLARE ABSTRACT FUNCTION GoBack() AS HRESULT
	DECLARE ABSTRACT FUNCTION GoForward() AS HRESULT
	DECLARE ABSTRACT FUNCTION GoHome() AS HRESULT
	DECLARE ABSTRACT FUNCTION GoSearch() AS HRESULT
	DECLARE ABSTRACT FUNCTION Navigate(BYVAL URL AS AFX_BSTR, BYVAL Flags AS VARIANT PTR, BYVAL TargetFrameName AS VARIANT PTR, BYVAL PostData AS VARIANT PTR, BYVAL Headers AS VARIANT PTR) AS HRESULT
	DECLARE ABSTRACT FUNCTION Refresh() AS HRESULT
	DECLARE ABSTRACT FUNCTION Refresh2(BYVAL Level AS VARIANT PTR) AS HRESULT
	DECLARE ABSTRACT FUNCTION Stop() AS HRESULT
	DECLARE ABSTRACT FUNCTION get_Application(BYVAL ppDisp AS Afx_IDispatch PTR PTR) AS HRESULT
	DECLARE ABSTRACT FUNCTION get_Parent(BYVAL ppDisp AS Afx_IDispatch PTR PTR) AS HRESULT
	DECLARE ABSTRACT FUNCTION get_Container(BYVAL ppDisp AS Afx_IDispatch PTR PTR) AS HRESULT
	DECLARE ABSTRACT FUNCTION get_Document(BYVAL ppDisp AS Afx_IDispatch PTR PTR) AS HRESULT
	DECLARE ABSTRACT FUNCTION get_TopLevelContainer(BYVAL pBool AS VARIANT_BOOL PTR) AS HRESULT
	DECLARE ABSTRACT FUNCTION get_Type(BYVAL Type AS AFX_BSTR PTR) AS HRESULT
	DECLARE ABSTRACT FUNCTION get_Left(BYVAL pl AS LONG PTR) AS HRESULT
	DECLARE ABSTRACT FUNCTION put_Left(BYVAL Left AS LONG) AS HRESULT
	DECLARE ABSTRACT FUNCTION get_Top(BYVAL pl AS LONG PTR) AS HRESULT
	DECLARE ABSTRACT FUNCTION put_Top(BYVAL Top AS LONG) AS HRESULT
	DECLARE ABSTRACT FUNCTION get_Width(BYVAL pl AS LONG PTR) AS HRESULT
	DECLARE ABSTRACT FUNCTION put_Width(BYVAL Width AS LONG) AS HRESULT
	DECLARE ABSTRACT FUNCTION get_Height(BYVAL pl AS LONG PTR) AS HRESULT
	DECLARE ABSTRACT FUNCTION put_Height(BYVAL Height AS LONG) AS HRESULT
	DECLARE ABSTRACT FUNCTION get_LocationName(BYVAL LocationName AS AFX_BSTR PTR) AS HRESULT
	DECLARE ABSTRACT FUNCTION get_LocationURL(BYVAL LocationURL AS AFX_BSTR PTR) AS HRESULT
	DECLARE ABSTRACT FUNCTION get_Busy(BYVAL pBool AS VARIANT_BOOL PTR) AS HRESULT
END TYPE

TYPE AFX_LPIWEBBROWSER AS Afx_IWebBrowser PTR

' ########################################################################################
' Interface name = IWebBrowserApp
' IID = {0002DF05-0000-0000-C000-000000000046}
' Inherited interface = IWebBrowser
' Deprecated. Use IWebBrowser2.
' ########################################################################################

#define __Afx_IWebBrowserApp_INTERFACE_DEFINED__
TYPE Afx_IWebBrowserApp AS Afx_IWebBrowserApp_

TYPE Afx_IWebBrowserApp_ EXTENDS Afx_IWebBrowser
	DECLARE ABSTRACT FUNCTION Quit () AS HRESULT
	DECLARE ABSTRACT FUNCTION ClientToWindow(BYVAL pcx AS LONG PTR, BYVAL pcy AS LONG PTR) AS HRESULT
	DECLARE ABSTRACT FUNCTION PutProperty(BYVAL Property AS AFX_BSTR, BYVAL vtValue AS VARIANT) AS HRESULT
	DECLARE ABSTRACT FUNCTION GetProperty(BYVAL Property AS AFX_BSTR, BYVAL pvtValue AS VARIANT PTR) AS HRESULT
	DECLARE ABSTRACT FUNCTION get_Name(BYVAL Name AS AFX_BSTR PTR) AS HRESULT
	DECLARE ABSTRACT FUNCTION get_HWND(BYVAL pHWND AS SHANDLE_PTR PTR) AS HRESULT
	DECLARE ABSTRACT FUNCTION get_FullName(BYVAL FullName AS AFX_BSTR PTR) AS HRESULT
	DECLARE ABSTRACT FUNCTION get_Path(BYVAL Path AS AFX_BSTR PTR) AS HRESULT
	DECLARE ABSTRACT FUNCTION get_Visible(BYVAL pBool AS VARIANT_BOOL PTR) AS HRESULT
	DECLARE ABSTRACT FUNCTION put_Visible(BYVAL Value AS VARIANT_BOOL) AS HRESULT
	DECLARE ABSTRACT FUNCTION get_StatusBar(BYVAL pBool AS VARIANT_BOOL PTR) AS HRESULT
	DECLARE ABSTRACT FUNCTION put_StatusBar(BYVAL Value AS VARIANT_BOOL) AS HRESULT
	DECLARE ABSTRACT FUNCTION get_StatusText(BYVAL StatusText AS AFX_BSTR PTR) AS HRESULT
	DECLARE ABSTRACT FUNCTION put_StatusText(BYVAL StatusText AS AFX_BSTR) AS HRESULT
	DECLARE ABSTRACT FUNCTION get_ToolBar(BYVAL Value AS LONG PTR) AS HRESULT
	DECLARE ABSTRACT FUNCTION put_ToolBar(BYVAL Value AS LONG) AS HRESULT
	DECLARE ABSTRACT FUNCTION get_MenuBar(BYVAL Value AS VARIANT_BOOL PTR) AS HRESULT
	DECLARE ABSTRACT FUNCTION put_MenuBar(BYVAL Value AS VARIANT_BOOL) AS HRESULT
	DECLARE ABSTRACT FUNCTION get_FullScreen(BYVAL pbFullScreen AS VARIANT_BOOL PTR) AS HRESULT
	DECLARE ABSTRACT FUNCTION put_FullScreen(BYVAL bFullScreen AS VARIANT_BOOL) AS HRESULT
END TYPE

TYPE AFX_LPIWEBBROWSERAPP AS Afx_IWebBrowserApp PTR

' ########################################################################################
' Interface name = IWebBrowser2
' IID = {D30C1661-CDAF-11D0-8A3E-00C04FC9E26E}
' Inherited interface = IWebBrowserApp
' This interface enables applications to implement an instance of the WebBrowser control
' (Microsoft ActiveX control) or control an instance of the InternetExplorer application
' (OLE Automation). Note that not all of the methods listed below are supported by the
' WebBrowser control.
' The IWebBrowser2 interface derives from IDispatch indirectly. IWebBrowser2 derives from
' IWebBrowserApp, which in turn derives from IWebBrowser, which finally derives from
' IDispatch. The IWebBrowser and IWebBrowserApp interfaces are deprecated.
' ########################################################################################

#define __Afx_IWebBrowser2_INTERFACE_DEFINED__
TYPE Afx_IWebBrowser2 AS Afx_IWebBrowser2_

TYPE Afx_IWebBrowser2_ EXTENDS Afx_IWebBrowserApp
   DECLARE ABSTRACT FUNCTION Navigate2 (BYVAL URL AS VARIANT PTR, BYVAL Flags AS VARIANT PTR = NULL, BYVAL TargetFrameName AS VARIANT PTR = NULL, BYVAL PostData AS VARIANT PTR = NULL, BYVAL Headers AS VARIANT PTR = NULL) AS HRESULT
   DECLARE ABSTRACT FUNCTION QueryStatusWB (BYVAL cmdID AS OLECMDID, BYVAL pcmdf AS OLECMDF PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION ExecWB (BYVAL cmdID AS OLECMDID, BYVAL cmdexecopt AS OLECMDEXECOPT, BYVAL pvaIn AS VARIANT PTR, BYVAL pvaOut AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION ShowBrowserBar (BYVAL pvaClsid AS VARIANT PTR, BYVAL pvarShow AS VARIANT PTR = NULL, BYVAL pvarSize AS VARIANT PTR = NULL) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ReadyState (BYVAL plReadyState AS READYSTATE PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Offline (BYVAL pbOffline AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Offline (BYVAL bOffline AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Silent (BYVAL pbSilent AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Silent (BYVAL bSilent AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_RegisterAsBrowser (BYVAL pbRegister AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_RegisterAsBrowser (BYVAL bRegister AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_RegisterAsDropTarget (BYVAL pbRegister AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_RegisterAsDropTarget (BYVAL bRegister AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_TheaterMode (BYVAL pbRegister AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_TheaterMode (BYVAL bRegister AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_AddressBar (BYVAL Value AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_AddressBar (BYVAL Value AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Resizable (BYVAL Value AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Resizable (BYVAL Value AS VARIANT_BOOL) AS HRESULT
END TYPE

TYPE AFX_LPIWEBBROWSER2 AS Afx_IWebBrowser2 PTR

' ########################################################################################
' Interface name = IShellWindows
' IID = {85CB6900-4D95-11CF-960C-0080C7F4EE85}
' Inherited interface = IDispatch
' ########################################################################################

#define __Afx_IShellWindows_INTERFACE_DEFINED__
TYPE Afx_IShellWindows AS Afx_IShellWindows_

TYPE Afx_IShellWindows_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_Count(BYVAL Count AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Item(BYVAL index AS VARIANT, BYVAL Folder AS Afx_IDispatch PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION _NewEnum(BYVAL ppunk AS Afx_IUnknown PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Register(BYVAL pid AS Afx_IDispatch PTR, BYVAL hwnd AS Long, BYVAL swClass AS long, BYVAL plCookie AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION RegisterPending(BYVAL lThreadId AS long, BYVAL pvarloc AS VARIANT PTR, BYVAL pvarlocRoot AS VARIANT PTR, BYVAL swClass AS long, BYVAL plCookie AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Revoke(BYVAL lCookie AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION OnNavigate(BYVAL lCookie AS long, BYVAL pvarLoc AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION OnActivated(BYVAL lCookie AS long, BYVAL fActive AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION FindWindowSW(BYVAL pvarLoc AS VARIANT PTR, BYVAL pvarLocRoot AS VARIANT PTR, BYVAL swClass AS long, BYVAL phwnd AS LONG PTR, BYVAL swfwOptions AS long, BYVAL ppdispOut AS Afx_IDispatch PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION OnCreated(BYVAL lCookie AS long, BYVAL punk AS Afx_IUnknown PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION ProcessAttachDetach(BYVAL fAttach AS VARIANT_BOOL) AS HRESULT
END TYPE

TYPE AFX_LPISHELLWINDOWS AS Afx_IShellWindows PTR

' ########################################################################################
' Interface name = IShellUIHelper
' IID = 729FE2F8-1EA8-11D1-8F85-00C04FC2FBE1
' Inherited interface = IDispatch
' This interface provides access to features available in the Microsoft Windows Shell API.
' ########################################################################################

#define __Afx_IShellUIHelper_INTERFACE_DEFINED__
TYPE Afx_IShellUIHelper AS Afx_IShellUIHelper_

TYPE Afx_IShellUIHelper_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION ResetFirstBootMode() AS HRESULT
   DECLARE ABSTRACT FUNCTION ResetSafeMode() AS HRESULT
   DECLARE ABSTRACT FUNCTION RefreshOfflineDesktop() AS HRESULT
   DECLARE ABSTRACT FUNCTION AddFavorite(BYVAL URL AS AFX_BSTR, BYVAL Title AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION AddChannel(BYVAL URL AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION AddDesktopComponent(BYVAL URL AS AFX_BSTR, BYVAL Type AS AFX_BSTR, BYVAL Left AS VARIANT PTR, BYVAL Top AS VARIANT PTR, BYVAL Width AS VARIANT PTR, BYVAL Height AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION IsSubscribed(BYVAL URL AS AFX_BSTR, BYVAL pBool AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION NavigateAndFind(BYVAL URL AS AFX_BSTR, BYVAL strQuery AS AFX_BSTR, BYVAL varTargetFrame AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION ImportExportFavorites(BYVAL fImport AS VARIANT_BOOL, BYVAL strImpExpPath AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION AutoCompleteSaveForm(BYVAL Form AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION AutoScan(BYVAL strSearch AS AFX_BSTR, BYVAL strFailureUrl AS AFX_BSTR, BYVAL pvarTargetFrame AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION AutoCompleteAttach(BYVAL Reserved AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION ShowBrowserUI(BYVAL bstrName AS AFX_BSTR, BYVAL pvarIn AS VARIANT PTR, BYVAL pvarOut AS VARIANT PTR) AS HRESULT
END TYPE

TYPE AFX_LPISHELLUIHELPER AS Afx_IShellUIHelper PTR

' ########################################################################################
' Interface name = IShellUIHelper2 [IE7]
' IID = {A7FE6EDA-1932-4281-B881-87B31B8BC52C}
' Attributes = 4160 [&H1040] [Dual] [Dispatchable]
' Inherited interface = IShellUIHelper
' ########################################################################################

#define __Afx_IShellUIHelper2_INTERFACE_DEFINED__
TYPE Afx_IShellUIHelper2 AS Afx_IShellUIHelper2_

TYPE Afx_IShellUIHelper2_ EXTENDS Afx_IShellUIHelper
   DECLARE ABSTRACT FUNCTION AddSearchProvider (BYVAL URL AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION RunOnceShown () AS HRESULT
   DECLARE ABSTRACT FUNCTION SkipRunOnce () AS HRESULT
   DECLARE ABSTRACT FUNCTION CustomizeSettings (BYVAL fSQM AS VARIANT_BOOL, BYVAL fPhishing AS VARIANT_BOOL, BYVAL bstrLocale AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION SqmEnabled (BYVAL pfEnabled AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION PhishingEnabled (BYVAL pfEnabled AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION BrandImageUri (BYVAL pbstrUri AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION SkipTabsWelcome () AS HRESULT
   DECLARE ABSTRACT FUNCTION DiagnoseConnection () AS HRESULT
   DECLARE ABSTRACT FUNCTION CustomizeClearType (BYVAL fSet AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION IsSearchProviderInstalled (BYVAL URL AS AFX_BSTR, BYVAL pdwResult AS DWORD) AS HRESULT
   DECLARE ABSTRACT FUNCTION IsSearchMigrated (BYVAL pfMigrated AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION DefaultSearchProvider (BYVAL pbstrName AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION RunOnceRequiredSettingsComplete (BYVAL fComplete AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION RunOnceHasShown (BYVAL pfShown AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION SearchGuideUrl (BYVAL pbstrUrl AS AFX_BSTR PTR) AS HRESULT
END TYPE

TYPE AFX_LPISHELLUIHELPER2 AS Afx_IShellUIHelper2 PTR

' ########################################################################################
' Interface name = IShellUIHelper3 [IE8]
' IID = {528DF2EC-D419-40BC-9B6D-DCDBF9C1B25D}
' Attributes = 4160 [&H1040] [Dual] [Dispatchable]
' Inherited interface = IShellUIHelper2
' ########################################################################################

#define __Afx_IShellUIHelper3_INTERFACE_DEFINED__
TYPE Afx_IShellUIHelper3 AS Afx_IShellUIHelper3_

TYPE Afx_IShellUIHelper3_ EXTENDS Afx_IShellUIHelper2
   DECLARE ABSTRACT FUNCTION AddService (BYVAL URL AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION IsServiceInstalled (BYVAL URL AS AFX_BSTR, BYVAL Verb AS AFX_BSTR, BYVAL pdwResult AS DWORD) AS HRESULT
   DECLARE ABSTRACT FUNCTION InPrivateFilteringEnabled (BYVAL pfEnabled AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION AddToFavoritesBar (BYVAL URL AS AFX_BSTR, BYVAL Title AS AFX_BSTR, BYVAL Type AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION BuildNewTabPage () AS HRESULT
   DECLARE ABSTRACT FUNCTION SetRecentlyClosedVisible (BYVAL fVisible AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION SetActivitiesVisible (BYVAL fVisible AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION ContentDiscoveryReset () AS HRESULT
   DECLARE ABSTRACT FUNCTION IsSuggestedSitesEnabled (BYVAL pfEnabled AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION EnableSuggestedSites (BYVAL fEnable AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION NavigateToSuggestedSites (BYVAL bstrRelativeUrl AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION ShowTabsHelp () AS HRESULT
   DECLARE ABSTRACT FUNCTION ShowInPrivateHelp () AS HRESULT
END TYPE

TYPE AFX_LPISHELLUIHELPER3 AS Afx_IShellUIHelper3 PTR

' ########################################################################################
' Interface name = IShellFavoritesNameSpace
' IID = {55136804-B2DE-11D1-B9F2-00A0C98BC547}
' Attributes = 4176 [&H1050] [Hidden] [Dual] [Dispatchable]
' Inherited interface = IDispatch
' ########################################################################################

#define __Afx_IShellFavoritesNameSpace_INTERFACE_DEFINED__
TYPE Afx_IShellFavoritesNameSpace AS Afx_IShellFavoritesNameSpace_

TYPE Afx_IShellFavoritesNameSpace_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION MoveSelectionUp() AS HRESULT
   DECLARE ABSTRACT FUNCTION MoveSelectionDown() AS HRESULT
   DECLARE ABSTRACT FUNCTION ResetSort() AS HRESULT
   DECLARE ABSTRACT FUNCTION NewFolder() AS HRESULT
   DECLARE ABSTRACT FUNCTION Synchronize_() AS HRESULT
   DECLARE ABSTRACT FUNCTION Import_() AS HRESULT
   DECLARE ABSTRACT FUNCTION Export_() AS HRESULT
   DECLARE ABSTRACT FUNCTION InvokeContextMenuCommand(BYVAL strCommand AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION MoveSelectionTo() AS HRESULT
   DECLARE ABSTRACT FUNCTION get_SubscriptionsEnabled(BYVAL pBool AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION CreateSubscriptionForSelection(BYVAL pBool AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION DeleteSubscriptionForSelection(BYVAL pBool AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION SetRoot(BYVAL bstrFullPath AS AFX_BSTR) AS HRESULT
END TYPE

TYPE AFX_LPISHELLFAVORITESNAMESPACE AS Afx_IShellFavoritesNameSpace PTR

' ########################################################################################
' Interface name = IShellNameSpace
' IID = {E572D3C9-37BE-4AE2-825D-D521763E3108}
' Attributes = 4176 [&H1050] [Hidden] [Dual] [Dispatchable]
' Inherited interface = IShellFavoritesNameSpace
' ########################################################################################

#define __Afx_IShellNameSpace_INTERFACE_DEFINED__
TYPE Afx_IShellNameSpace AS Afx_IShellFavoritesNameSpace_

TYPE Afx_IShellNameSpace_ EXTENDS Afx_IShellFavoritesNameSpace
   DECLARE ABSTRACT FUNCTION get_EnumOptions(BYVAL pgrfEnumFlags AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_EnumOptions(BYVAL lVal AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_SelectedItem(BYVAL pItem AS Afx_IDispatch PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_SelectedItem(BYVAL pItem AS Afx_IDispatch PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Root(BYVAL pvar AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Root(BYVAL var AS VARIANT) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Depth(BYVAL piDepth AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Depth(BYVAL iDepth AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Mode(BYVAL puMode AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Mode(BYVAL uMode AS UINT) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Flags(BYVAL pdwFlags AS DWORD PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Flags(BYVAL dwFlags AS DWORD) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_TVFlags(BYVAL dwFlags AS DWORD) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_TVFlags(BYVAL dwFlags AS DWORD PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Columns(BYVAL bstrColumns AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Columns(BYVAL bstrColumns AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_CountViewTypes(BYVAL piTypes AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION SetViewType(BYVAL iType AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION SelectedItems(BYVAL ppid AS Afx_IDispatch PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Expand(BYVAL var AS VARIANT, BYVAL iDepth AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION UnselectAll() AS HRESULT
END TYPE

TYPE AFX_LPISHELLNAMESPACE AS Afx_IShellNameSpace PTR

' ########################################################################################
' Interface name = IScriptErrorList
' IID = {F3470F24-15FD-11D2-BB2E-00805FF7EFCA}
' Attributes = 4176 [&H1050] [Hidden] [Dual] [Dispatchable]
' Inherited interface = IDispatch
' ########################################################################################

#define __Afx_IScriptErrorList_INTERFACE_DEFINED__
TYPE Afx_IScriptErrorList AS Afx_IScriptErrorList_

TYPE Afx_IScriptErrorList_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION advanceError() AS HRESULT
   DECLARE ABSTRACT FUNCTION retreatError() AS HRESULT
   DECLARE ABSTRACT FUNCTION canAdvanceError(BYVAL pfCanAdvance AS WINBOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION canRetreatError(BYVAL pfCanRetreat AS WINBOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION getErrorLine(BYVAL plLine AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION getErrorChar(BYVAL plChar AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION getErrorCode(BYVAL plCode AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION getErrorMsg(BYVAL pstr AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION getErrorUrl(BYVAL pstr AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION getAlwaysShowLockState(BYVAL pfAlwaysShowLocked AS WINBOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION getDetailsPaneOpen(BYVAL pfDetailsPaneOpen AS WINBOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION setDetailsPaneOpen(BYVAL fDetailsPaneOpen AS WINBOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION getPerErrorDisplay(BYVAL pfPerErrorDisplay AS WINBOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION setPerErrorDisplay(BYVAL fPerErrorDisplay AS WINBOOL) AS HRESULT
END TYPE

TYPE AFX_LPISCRIPTERRORLIST AS Afx_IScriptErrorList PTR

END NAMESPACE

' ========================================================================================
' Adds a favorite. Use for backward compatibility only.
' ========================================================================================
PRIVATE FUNCTION AddUrlToFavorites (BYVAL hwnd AS HWND, BYVAL pszUrlW AS WSTRING PTR, BYVAL pszTitleW AS WSTRING PTR, BYVAL fDisplayUI AS LONG) AS HRESULT
   DIM AS ANY PTR pLib = DyLibLoad("shdocvw.dll")
   IF pLib = 0 THEN EXIT FUNCTION
   DIM pProc AS FUNCTION (BYVAL hwnd AS HWND, BYVAL pszUrlW AS WSTRING PTR, BYVAL pszTitleW AS WSTRING PTR, BYVAL fDisplayUI AS LONG) AS HRESULT
   pProc = DyLibSymbol(pLib, "AddUrlToFavorites")
   IF pProc = 0 THEN EXIT FUNCTION
   FUNCTION = pProc(hwnd, pszUrlW, pszTitleW, fDisplayUI)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Displays the standard Microsoft Internet Explorer Add Favorite dialog box.
' ========================================================================================
PRIVATE FUNCTION DoAddToFavDlg (BYVAL hwnd AS HWND, BYVAL pszInitDir AS ZSTRING PTR, BYVAL cchInitDir AS UINT, BYVAL pszFile AS ZSTRING PTR, BYVAL cchFile AS UINT, BYVAL pidlBrowse AS LPITEMIDLIST) AS LONG
   DIM AS ANY PTR pLib = DyLibLoad("shdocvw.dll")
   IF pLib = 0 THEN EXIT FUNCTION
   DIM pProc AS FUNCTION (BYVAL hwnd AS HWND, BYVAL pszInitDir AS ZSTRING PTR, BYVAL cchInitDir AS UINT, BYVAL pszFile AS ZSTRING PTR, BYVAL cchFile AS UINT, BYVAL pidlBrowse AS LPITEMIDLIST) AS LONG
   pProc = DyLibSymbol(pLib, "DoAddToFavDlg")
   IF pProc = 0 THEN EXIT FUNCTION
   FUNCTION = pProc(hwnd, pszInitDir, cchInitDir, pszFile, cchInitDir, pidlBrowse)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Displays the standard Microsoft Internet Explorer Organize Favorites dialog box.
' ========================================================================================
PRIVATE FUNCTION DoOrganizeFavDlg (BYVAL hwnd AS HWND, BYVAL pszInitDir AS ZSTRING PTR) AS LONG
   DIM AS ANY PTR pLib = DyLibLoad("shdocvw.dll")
   IF pLib = 0 THEN EXIT FUNCTION
   DIM pProc AS FUNCTION (BYVAL hwnd AS HWND, BYVAL pszInitDir AS ZSTRING PTR) AS LONG
   pProc = DyLibSymbol(pLib, "DoOrganizeFavDlg")
   IF pProc = 0 THEN EXIT FUNCTION
   FUNCTION = pProc(hwnd, pszInitDir)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================


' ########################################################################################
'                          *** Associated Browser Functions ***
' ########################################################################################

' ========================================================================================
' Launches the Internet Properties dialog box with the Connections tab displayed.
' ========================================================================================
PRIVATE FUNCTION LaunchConnectionDialog (BYVAL hParent AS HWND) AS LONG
   DIM AS ANY PTR pLib = DyLibLoad("inetcpl.cpl")
   IF pLib = 0 THEN EXIT FUNCTION
   DIM pProc AS FUNCTION (BYVAL hParent AS HWND) AS LONG
   pProc = DyLibSymbol(pLib, "LaunchConnectionDialog")
   IF pProc = 0 THEN EXIT FUNCTION
   FUNCTION = pProc(hParent)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Launches the Internet Properties dialog box with the General tab displayed.
' ========================================================================================
PRIVATE FUNCTION LaunchInternetControlPanel (BYVAL hParent AS HWND) AS LONG
   DIM AS ANY PTR pLib = DyLibLoad("inetcpl.cpl")
   IF pLib = 0 THEN EXIT FUNCTION
   DIM pProc AS FUNCTION (BYVAL hParent AS HWND) AS LONG
   pProc = DyLibSymbol(pLib, "LaunchInternetControlPanel")
   IF pProc = 0 THEN EXIT FUNCTION
   FUNCTION = pProc(hParent)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Launches the Internet Properties dialog box with the Privacy tab displayed.
' ========================================================================================
PRIVATE FUNCTION LaunchPrivacyDialog (BYVAL hParent AS HWND) AS LONG
   DIM AS ANY PTR pLib = DyLibLoad("inetcpl.cpl")
   IF pLib = 0 THEN EXIT FUNCTION
   DIM pProc AS FUNCTION (BYVAL hParent AS HWND) AS LONG
   pProc = DyLibSymbol(pLib, "LaunchPrivacyDialog")
   IF pProc = 0 THEN EXIT FUNCTION
   FUNCTION = pProc(hParent)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Launches the Internet Properties dialog box with the Security tab displayed.
' ========================================================================================
PRIVATE FUNCTION LaunchSecurityDialog (BYVAL hParent AS HWND) AS LONG
   DIM AS ANY PTR pLib = DyLibLoad("inetcpl.cpl")
   IF pLib = 0 THEN EXIT FUNCTION
   DIM pProc AS FUNCTION (BYVAL hParent AS HWND) AS LONG
   pProc = DyLibSymbol(pLib, "LaunchSecurityDialog")
   IF pProc = 0 THEN EXIT FUNCTION
   FUNCTION = pProc(hParent)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================
