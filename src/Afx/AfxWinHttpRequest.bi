' ########################################################################################
' Microsoft Windows
' File: AfxWinHttpRequest.bi
' Library name: WinHttp
' Version: 5.1
' Documentation string: Microsoft WinHTTP Services, version 5.1
' Path: C:\WINDOWS\system32\WINHTTP.dll
' Library GUID: {662901FC-6951-4854-9EB2-D9A2570F2B2E}
' Portions Copyright (c) Microsoft Corporation.
' Compiler: Free Basic 32 & 64 bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "win/ole2.bi"
#include once "win/unknwnbase.bi"

NAMESPACE Afx

' // The definition for BSTR in the FreeBASIC headers was inconveniently changed to WCHAR
#ifndef AFX_BSTR
   #define AFX_BSTR WSTRING PTR
#endif

' // From httprequestid.h
CONST DISPID_HTTPREQUEST_BASE                     = &H00000001
CONST DISPID_HTTPREQUEST_OPEN                     = DISPID_HTTPREQUEST_BASE
CONST DISPID_HTTPREQUEST_SETREQUESTHEADER         = DISPID_HTTPREQUEST_BASE + 1
CONST DISPID_HTTPREQUEST_GETRESPONSEHEADER        = DISPID_HTTPREQUEST_BASE + 2
CONST DISPID_HTTPREQUEST_GETALLRESPONSEHEADERS    = DISPID_HTTPREQUEST_BASE + 3
CONST DISPID_HTTPREQUEST_SEND                     = DISPID_HTTPREQUEST_BASE + 4
CONST DISPID_HTTPREQUEST_OPTION                   = DISPID_HTTPREQUEST_BASE + 5
CONST DISPID_HTTPREQUEST_STATUS                   = DISPID_HTTPREQUEST_BASE + 6
CONST DISPID_HTTPREQUEST_STATUSTEXT               = DISPID_HTTPREQUEST_BASE + 7
CONST DISPID_HTTPREQUEST_RESPONSETEXT             = DISPID_HTTPREQUEST_BASE + 8
CONST DISPID_HTTPREQUEST_RESPONSEBODY             = DISPID_HTTPREQUEST_BASE + 9
CONST DISPID_HTTPREQUEST_RESPONSESTREAM           = DISPID_HTTPREQUEST_BASE + 10
CONST DISPID_HTTPREQUEST_ABORT                    = DISPID_HTTPREQUEST_BASE + 11
CONST DISPID_HTTPREQUEST_SETPROXY                 = DISPID_HTTPREQUEST_BASE + 12
CONST DISPID_HTTPREQUEST_SETCREDENTIALS           = DISPID_HTTPREQUEST_BASE + 13
CONST DISPID_HTTPREQUEST_WAITFORRESPONSE          = DISPID_HTTPREQUEST_BASE + 14
CONST DISPID_HTTPREQUEST_SETTIMEOUTS              = DISPID_HTTPREQUEST_BASE + 15
CONST DISPID_HTTPREQUEST_SETCLIENTCERTIFICATE     = DISPID_HTTPREQUEST_BASE + 16
CONST DISPID_HTTPREQUEST_SETAUTOLOGONPOLICY       = DISPID_HTTPREQUEST_BASE + 17

' ========================================================================================
' ProgIDs (Program identifiers)
' ========================================================================================

CONST AFX_PROGID_WinHttpRequest51 = "WinHttp.WinHttpRequest.5.1"

' ========================================================================================
' ClsIDs (Class identifiers)
' ========================================================================================

CONST AFX_CLSID_WinHttpRequest = "{2087C2F4-2CEF-4953-A8AB-66779B670495}"

' ========================================================================================
' IIDs (Interface identifiers)
' ========================================================================================

CONST AFX_IID_IWinHttpRequest = "{016FE2EC-B2C8-45F8-B23B-39E53A75396B}"
CONST AFX_IID_IWinHttpRequestEvents = "{F97F4E15-B787-4212-80D1-D380CBBF982E}"

TYPE HTTPREQUEST_PROXY_SETTING AS LONG
CONST HTTPREQUEST_PROXYSETTING_DEFAULT      = &h00000000
CONST HTTPREQUEST_PROXYSETTING_PRECONFIG    = &h00000000
CONST HTTPREQUEST_PROXYSETTING_DIRECT       = &h00000001
CONST HTTPREQUEST_PROXYSETTING_PROXY        = &h00000002

TYPE HTTPREQUEST_SETCREDENTIALS_FLAGS AS LONG
CONST HTTPREQUEST_SETCREDENTIALS_FOR_SERVER = &h00000000
CONST HTTPREQUEST_SETCREDENTIALS_FOR_PROXY  = &h00000001

' ========================================================================================
' WinHttpRequestOption enum
' IID: {12782009-FE90-4877-9730-E5E183669B19}
' WinHttpRequest Options
' ========================================================================================

enum WinHttpRequestOption
   WinHttpRequestOption_UserAgentString = 0
   WinHttpRequestOption_URL = 1
   WinHttpRequestOption_URLCodePage = 2
   WinHttpRequestOption_EscapePercentInURL = 3
   WinHttpRequestOption_SslErrorIgnoreFlags = 4
   WinHttpRequestOption_SelectCertificate = 5
   WinHttpRequestOption_EnableRedirects = 6
   WinHttpRequestOption_UrlEscapeDisable = 7
   WinHttpRequestOption_UrlEscapeDisableQuery = 8
   WinHttpRequestOption_SecureProtocols = 9
   WinHttpRequestOption_EnableTracing = 10
   WinHttpRequestOption_RevertImpersonationOverSsl = 11
   WinHttpRequestOption_EnableHttpsToHttpRedirects = 12
   WinHttpRequestOption_EnablePassportAuthentication = 13
   WinHttpRequestOption_MaxAutomaticRedirects = 14
   WinHttpRequestOption_MaxResponseHeaderSize = 15
   WinHttpRequestOption_MaxResponseDrainSize = 16
   WinHttpRequestOption_EnableHttp1_1 = 17
   WinHttpRequestOption_EnableCertificateRevocationCheck = 18
   WinHttpRequestOption_RejectUserpwd = 19
end enum

' ========================================================================================
' WinHttpRequestAutoLogonPolicy enum
' IID: {9D8A6DF8-13DE-4B1F-A330-67C719D62514}
' ========================================================================================

enum WinHttpRequestAutoLogonPolicy
   AutoLogonPolicy_Always = 0
   AutoLogonPolicy_OnlyIfBypassProxy = 1
   AutoLogonPolicy_Never = 2
end enum

' ========================================================================================
' WinHttpRequestSslErrorFlags enum
' IID: {152A1CA2-55A9-43A3-B187-0605BB886349}
' ========================================================================================

enum WinHttpRequestSslErrorFlags
   SslErrorFlag_UnknownCA = 256
   SslErrorFlag_CertWrongUsage = 512
   SslErrorFlag_CertCNInvalid = 4096
   SslErrorFlag_CertDateInvalid = 8192
   SslErrorFlag_Ignore_All = 13056
end enum

' ========================================================================================
' WinHttpRequestSecureProtocols enum
' IID: {6B2C51C1-A8EA-46BD-B928-C9B76F9F14DD}
' ========================================================================================

enum WinHttpRequestSecureProtocols
   SecureProtocol_SSL2 = 8
   SecureProtocol_SSL3 = 32
   SecureProtocol_TLS1 = 128
   SecureProtocol_ALL = 168
end enum

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

TYPE Afx_IWinHttpRequest AS Afx_IWinHttpRequest_
TYPE Afx_IWinHttpRequestEvents AS Afx_IWinHttpRequestEvents_

' ########################################################################################
' Interface name = IWinHttpRequest
' IID = {016FE2EC-B2C8-45F8-B23B-39E53A75396B}
' IWinHttpRequest Interface
' Attributes = 4288 [&H10C0] [Dual] [Nonextensible] [Dispatchable]
' Inherited interface = IDispatch
' ########################################################################################

#ifndef __Afx_IWinHttpRequest_INTERFACE_DEFINED__
#define __Afx_IWinHttpRequest_INTERFACE_DEFINED__

TYPE Afx_IWinHttpRequest_ EXTENDS Afx_IDispatch

   DECLARE ABSTRACT FUNCTION SetProxy (BYVAL ProxySetting AS HTTPREQUEST_PROXY_SETTING, BYVAL ProxyServer AS VARIANT = TYPE(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), BYVAL BypassList AS VARIANT = TYPE(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND)) AS HRESULT
   DECLARE ABSTRACT FUNCTION SetCredentials (BYVAL UserName AS AFX_BSTR, BYVAL Password AS AFX_BSTR, BYVAL Flags AS HTTPREQUEST_SETCREDENTIALS_FLAGS) AS HRESULT
   DECLARE ABSTRACT FUNCTION Open (BYVAL Method AS AFX_BSTR, BYVAL Url AS AFX_BSTR, BYVAL Async AS VARIANT = TYPE(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND)) AS HRESULT
   DECLARE ABSTRACT FUNCTION SetRequestHeader (BYVAL Header AS AFX_BSTR, BYVAL Value AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetResponseHeader (BYVAL Header AS AFX_BSTR, BYVAL Value AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetAllResponseHeaders (BYVAL Headers AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Send (BYVAL Body AS VARIANT = TYPE(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND)) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Status (BYVAL Status AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_StatusText (BYVAL Status AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ResponseText (BYVAL Body AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ResponseBody (BYVAL Body AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ResponseStream (BYVAL Body AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Option (BYVAL Option AS WinHttpRequestOption, BYVAL Value AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Option (BYVAL Option AS WinHttpRequestOption, BYVAL Value AS VARIANT) AS HRESULT
   DECLARE ABSTRACT FUNCTION WaitForResponse (BYVAL Timeout AS VARIANT = TYPE(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), BYVAL Succeeded AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Abort () AS HRESULT
   DECLARE ABSTRACT FUNCTION SetTimeouts (BYVAL ResolveTimeout AS LONG, BYVAL ConnectTimeout AS LONG, BYVAL SendTimeout AS LONG, BYVAL ReceiveTimeout AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION SetClientCertificate (BYVAL ClientCertificate AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION SetAutoLogonPolicy (BYVAL AutoLogonPolicy AS WinHttpRequestAutoLogonPolicy) AS HRESULT

END TYPE

TYPE AFX_LPWINHTTPREQUEST AS Afx_IWinHttpRequest_ PTR
#endif

' ########################################################################################
' Class CIWinHttpRequestEvents
' Interface name = IWinHttpRequestEvents
' IID = {F97F4E15-B787-4212-80D1-D380CBBF982E}
' IWinHttpRequestEvents Interface
' Attributes = 384 [&H180] [Nonextensible] [Oleautomation]
' ########################################################################################

#ifndef __Afx_IWinHttpRequestEvents_INTERFACE_DEFINED__
#define __Afx_IWinHttpRequestEvents_INTERFACE_DEFINED__

TYPE Afx_IWinHttpRequestEvents_ EXTENDS Afx_IUnknown

   DECLARE ABSTRACT FUNCTION OnResponseStart (BYVAL Status AS LONG, BYVAL ContentType AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION OnResponseDataAvailable (BYVAL Data AS SAFEARRAY PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION OnResponseFinished () AS HRESULT
   DECLARE ABSTRACT FUNCTION OnError (BYVAL ErrorNumber AS LONG, BYVAL ErrorDescription AS AFX_BSTR) AS HRESULT

END TYPE

TYPE AFX_LPWINHTTPREQUESTEVENTS AS Afx_IWinHttpRequestEvents_ PTR
#endif

END NAMESPACE
