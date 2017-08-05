' ########################################################################################
' Microsoft Windows
' File: CWmiCli.inc
' Contents: WMI class.
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2017 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "win/wbemcli.bi"
#include once "crt/win32/stdlib.bi"
#include once "Afx/AfxWmiDisp.bi"
#include once "Afx/AfxCOM.inc"

NAMESPACE Afx

' ========================================================================================
' Macro for debug
' To allow debugging, define _CWMICLI_DEBUG_ 1 in your application before including this file.
' ========================================================================================
#ifndef _CWMICLI_DEBUG_
   #define _CWMICLI_DEBUG_ 0
#ENDIF
#ifndef _CWMICLI_DP_
   #define _CWMICLI_DP_ 1
   #MACRO CWMICLI_DP(st)
      #IF (_CWMICLI_DEBUG_ = 1)
         OutputDebugStringW(st)
      #ENDIF
   #ENDMACRO
#ENDIF
' ========================================================================================

' ========================================================================================
' CWmiCli class
' ========================================================================================
TYPE CWmiCli

   Public:
      m_Result AS HRESULT                                  ' // HRESULT
      m_pIWbemLocator AS IWbemLocator PTR                  ' // IWbemLocator interface
      m_pIWbemServices AS IWbemServices PTR                ' // IWbemServices interface
      m_pIEnumWbemClassObject AS IEnumWbemClassObject PTR  ' // IEnumWbemClassObject collection
      m_pIWbemClassObject AS IWbemClassObject PTR          ' // IWbemClassObject interface
      m_pIWbemStatusCodeText AS IWbemStatusCodeText PTR    ' // IWbemStatusCodeText interface

   Private:
      m_bUninitCOM AS BOOLEAN

   Public:
      DECLARE CONSTRUCTOR (BYVAL pwszWmiNamespace AS WSTRING PTR)
      DECLARE DESTRUCTOR
      DECLARE FUNCTION GetLastResult () AS HRESULT
      DECLARE FUNCTION SetResult (BYVAL Result AS HRESULT) AS HRESULT
      DECLARE FUNCTION GetErrorCodeText (BYVAL hRes AS HRESULT = 0) AS CWSTR
      DECLARE FUNCTION GetFacilityCodeText (BYVAL hRes AS HRESULT = 0) AS CWSTR
      DECLARE FUNCTION ExecQuery (BYVAL pwszQuery AS WSTRING PTR) AS HRESULT
      DECLARE FUNCTION MoveNext () AS BOOLEAN
      DECLARE FUNCTION GetVar (BYVAL pwszName AS WSTRING PTR, BYVAL pvarRes AS VARIANT PTR) AS HRESULT
      DECLARE FUNCTION GetStr (BYVAL pwszName AS WSTRING PTR) AS CWSTR
      DECLARE FUNCTION WmiTimeToFileTime (BYVAL pwszDateTime AS WSTRING PTR, BYVAL bIsLocal AS BOOLEAN = -1) AS FILETIME
      DECLARE FUNCTION WmiTimeToSystemTime (BYVAL pwszDateTime AS WSTRING PTR, BYVAL bIsLocal AS BOOLEAN = -1) AS SYSTEMTIME

END TYPE
' ========================================================================================

' ========================================================================================
' CWmiCli constructor
' Note: This constructor only works with the local computer.
' ========================================================================================
PRIVATE CONSTRUCTOR CWmiCli (BYVAL pwszWmiNamespace AS WSTRING PTR)

   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE

   ' // Set the default security values for the process
   this.SetResult(CoInitializeSecurity( _
        NULL, _                          ' // PSECURITY_DESCRIPTOR
        -1, _                            ' // COM authentication
        NULL, _                          ' // Authentication services
        NULL, _                          ' // Reserved
        RPC_C_AUTHN_LEVEL_DEFAULT, _     ' // Default authentication
        RPC_C_IMP_LEVEL_IMPERSONATE, _   ' // Default Impersonation
        NULL, _                          ' // Authentication info
        EOAC_NONE, _                     ' // Additional capabilities
        NULL _                           ' // Reserved
        ))
   CWMICLI_DP("CWmiCli CONSTRUCTOR - CoInitializeSecurity Result = &H" & HEX(m_Result))

   ' // Create an instance of the IWbemLocator interface
   DIM CLSID_WbemLocator_ AS CLSID = (&h4590F811, &h1D3A, &h11D0, {&h89, &h1F, &h00, &hAA, &h00, &h4B, &h2E, &h24})
   DIM IID_IWbemLocator_  AS GUID  = (&hDC12A687, &h737F, &h11CF, {&h88, &h4D, &h00, &hAA, &h00, &h4B, &h2E, &h24})
   CoCreateInstance(@CLSID_WbemLocator_, NULL, CLSCTX_INPROC_SERVER, @IID_IWbemLocator_, @m_pIWbemLocator)
   CWMICLI_DP("CWmiCli CONSTRUCTOR - m_pIWbemLocator = " & WSTR(m_pIWbemLocator))
   IF m_pIWbemLocator = NULL THEN EXIT CONSTRUCTOR

   ' // Connect with the specified WMI namespace
   DIM bstrWmiNamespace AS AFX_BSTR
   IF pwszWmiNamespace THEN bstrWmiNamespace = SysAllocString(pwszWmiNamespace)
   this.SetResult(m_pIWbemLocator->lpvtbl->ConnectServer( _
        m_pIWbemLocator, _                   ' // Interface pointer
        bstrWmiNamespace, _                  ' // Object path of WMI namespace
        NULL, _                              ' // User name. NULL value = current user
        NULL, _                              ' // User password. NULL value = current
        NULL, _                              ' // Locale. NULL value indicates current
        WBEM_FLAG_CONNECT_USE_MAX_WAIT , _   ' // Security flags
        NULL, _                              ' // Authority (e.g. Kerberos)
        NULL, _                              ' // Context object
        @m_pIWbemServices _                  ' // pointer to IWbemServices proxy
        ))
   IF bstrWmiNamespace THEN SysFreeString(bstrWmiNamespace)
   CWMICLI_DP("CWmiCli CONSTRUCTOR - m_pIWbemServices = " & WSTR(m_pIWbemServices))
   IF m_pIWbemServices = NULL THEN EXIT CONSTRUCTOR

   ' // Set security levels on the proxy
   this.SetResult(CoSetProxyBlanket( _
        cast(IUnknown PTR, m_pIWbemServices), _   ' // Indicates the proxy to set
        RPC_C_AUTHN_WINNT, _                      ' // RPC_C_AUTHN_xxx
        RPC_C_AUTHZ_NONE, _                       ' // RPC_C_AUTHZ_xxx
        NULL, _                                   ' // Server principal name
        RPC_C_AUTHN_LEVEL_CALL, _                 ' // RPC_C_AUTHN_LEVEL_xxx
        RPC_C_IMP_LEVEL_IMPERSONATE, _            ' // RPC_C_IMP_LEVEL_xxx
        NULL, _                                   ' // client identity
        EOAC_NONE _                               ' // proxy capabilities
        ))
   CWMICLI_DP("CWmiCli CONSTRUCTOR - CoSetProxyBlanket - Result = &H" & HEX(m_Result))

   ' // Create an instance of the IWbemStatusCodeText interface
   DIM CLSID_WbemStatusCodeText_ AS CLSID = (&hEB87E1BD, &h3233, &h11D2, {&hAE, &hC9, &h00, &hC0, &h4F, &hB6, &h88, &h20})
   DIM IID_IWbemStatusCodeText_  AS GUID  = (&hEB87E1BC, &h3233, &h11D2, {&hAE, &hC9, &h00, &hC0, &h4F, &hB6, &h88, &h20})
   CoCreateInstance(@CLSID_WbemStatusCodeText_, NULL, CLSCTX_INPROC_SERVER, @IID_IWbemStatusCodeText_, @m_pIWbemStatusCodeText)
   CWMICLI_DP("CWmiCli CONSTRUCTOR - m_pIWbemStatusCodeText = " & WSTR(m_pIWbemStatusCodeText))

END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CWmiCli destructor
' ========================================================================================
PRIVATE DESTRUCTOR CWmiCli
   CWMICLI_DP("CWmiCli DESTRUCTOR")
   ' // Release the interfaces
   IF m_pIWbemLocator THEN m_pIWbemLocator->lpvtbl->Release(m_pIWbemLocator)
   IF m_pIWbemServices THEN m_pIWbemServices->lpvtbl->Release(m_pIWbemServices)
   IF m_pIEnumWbemClassObject THEN m_pIEnumWbemClassObject->lpvtbl->Release(m_pIEnumWbemClassObject)
   IF m_pIWbemClassObject THEN m_pIWbemClassObject->lpvtbl->Release(m_pIWbemClassObject)
   IF m_pIWbemStatusCodeText THEN m_pIWbemStatusCodeText->lpvtbl->Release(m_pIWbemStatusCodeText)
   ' // Free the COM library
   IF m_bUninitCOM THEN CoUninitialize
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the last result code.
' ========================================================================================
PRIVATE FUNCTION CWmiCli.GetLastResult () AS HRESULT
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last result code.
' ========================================================================================
PRIVATE FUNCTION CWmiCli.SetResult (BYVAL Result AS HRESULT) AS HRESULT
   m_Result = Result
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the text string description associated with the WMI error code.
' ========================================================================================
PRIVATE FUNCTION CWmiCli.GetErrorCodeText (BYVAL hRes AS HRESULT = 0) AS CWSTR
   CWMICLI_DP("CWmiCli.GetErrorCodeText")
   DIM cws AS CWSTR, bstrErrMsg AS AFX_BSTR
   IF m_pIWbemStatusCodeText = NULL THEN RETURN cws
   IF hRes = 0 THEN hRes = m_Result
   DIM hr AS HRESULT = m_pIWbemStatusCodeText->lpvtbl->GetErrorCodeText(m_pIWbemStatusCodeText, hRes, 0, 0, @bstrErrMsg)
   IF hr = WBEM_S_NO_ERROR AND bstrErrMsg <> NULL THEN cws = *bstrErrMsg
   IF bstrErrMsg THEN SysFreeString(bstrErrMsg)
   RETURN cws
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the name of the subsystem where the error occurred.
' ========================================================================================
PRIVATE FUNCTION CWmiCli.GetFacilityCodeText (BYVAL hRes AS HRESULT = 0) AS CWSTR
   CWMICLI_DP("CWmiCli.GetFacilityCodeText")
   DIM cws AS CWSTR, bstrErrMsg AS AFX_BSTR
   IF m_pIWbemStatusCodeText = NULL THEN RETURN cws
   IF hRes = 0 THEN hRes = m_Result
   DIM hr AS HRESULT = m_pIWbemStatusCodeText->lpvtbl->GetFacilityCodeText(m_pIWbemStatusCodeText, hRes, 0, 0, @bstrErrMsg)
   IF hr = WBEM_S_NO_ERROR AND bstrErrMsg <> NULL THEN cws = *bstrErrMsg
   IF bstrErrMsg THEN SysFreeString(bstrErrMsg)
   RETURN cws
END FUNCTION
' ========================================================================================

' ========================================================================================
' Executes a query
' ========================================================================================
PRIVATE FUNCTION CWmiCli.ExecQuery (BYVAL pwszQuery AS WSTRING PTR) AS HRESULT
   CWMICLI_DP("CWmiCli.ExecQuery")

   IF m_pIWbemServices = NULL THEN RETURN E_POINTER
   IF m_pIEnumWbemClassObject THEN
      m_pIEnumWbemClassObject->lpvtbl->Release(m_pIEnumWbemClassObject)
      m_pIEnumWbemClassObject = NULL
   END IF
   DIM bstrLanguage AS AFX_BSTR = SysAllocString("WQL")
   DIM bstrQuery AS AFX_BSTR = SysAllocString(pwszQuery)
   this.SetResult(m_pIWbemServices->lpvtbl->ExecQuery( _
      m_pIWbemServices, _                ' // Interface pointer
      bstrLanguage, _                    ' // Query language
      bstrQuery, _                       ' // Query string
      WBEM_FLAG_FORWARD_ONLY OR _        ' // Forward-only enumerator
      WBEM_FLAG_RETURN_IMMEDIATELY, _
      NULL, _                            ' // Typically NULL
      @m_pIEnumWbemClassObject))         ' // Pointer to the enumerator
   IF bstrLanguage THEN SysFreeString(bstrLanguage)
   IF bstrQuery THEN SysFreeString(bstrQuery)
   RETURN m_Result

END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the next object in the collection
' ========================================================================================
PRIVATE FUNCTION CWmiCli.MoveNext () AS BOOLEAN
   CWMICLI_DP("CWmiCli.MoveNext")

   DIM puReturned AS LONG
   this.SetResult(m_pIEnumWbemClassObject->lpvtbl->Next( _
         m_pIEnumWbemClassObject, _     ' // Interface pointer
         WBEM_INFINITE, _               ' // Timeout
         1, _                           ' // Number of requested objects
         @m_pIWbemClassObject, _        ' // Returned interface pointer
         @puReturned _                  ' // number of objects returned
         ))
   IF m_Result = WBEM_S_NO_ERROR THEN RETURN TRUE
   RETURN FALSE

END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the specified property value.
' ========================================================================================
PRIVATE FUNCTION CWmiCli.GetVar (BYVAL pwszName AS WSTRING PTR, BYVAL pvarRes AS VARIANT PTR) AS HRESULT
   CWMICLI_DP("CWmiCli.GetVar")

   IF pvarRes = NULL THEN RETURN E_POINTER
   VariantClear(pvarRes)
   DIM bstrName AS AFX_BSTR = SysAllocString(pwszName)
   this.SetResult(m_pIWbemClassObject->lpvtbl->Get( _
        m_pIWbemClassObject, _          ' // Interface pointer
        bstrName, _                     ' // Name of the desired property
        0, _                            ' // Reserved. Must be 0.
        pvarRes, _                      ' // Returned value for the qualifier
        NULL, _                         ' // CIM type of the property
        NULL _                          ' // Origin of the property
        ))
   IF bstrName THEN SysFreeString(bstrName)
   RETURN m_Result

END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the specified property value.
' ========================================================================================
PRIVATE FUNCTION CWmiCli.GetStr (BYVAL pwszName AS WSTRING PTR) AS CWSTR
   CWMICLI_DP("CWmiCli.GetStr")

   DIM vRes AS VARIANT, cws AS CWSTR
   DIM bstrName AS AFX_BSTR = SysAllocString(pwszName)
   this.SetResult(m_pIWbemClassObject->lpvtbl->Get( _
        m_pIWbemClassObject, _          ' // Interface pointer
        pwszName, _                     ' // Name of the desired property
        0, _                            ' // Reserved. Must be 0.
        @vRes, _                        ' // Returned value for the qualifier
        NULL, _                         ' // CIM type of the property
        NULL _                          ' // Origin of the property
        ))
   IF bstrName THEN SysFreeString(bstrName)
   RETURN AfxVarToStr(@vRes, TRUE)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts a date and time value in the CIM DATETIME format to the FILETIME format.
' ========================================================================================
PRIVATE FUNCTION CWmiCli.WmiTimeToFileTime (BYVAL pwszDateTime AS WSTRING PTR, BYVAL bIsLocal AS BOOLEAN = -1) AS FILETIME
   CWMICLI_DP("CWmiCli.WmiTimeToFileTime")

   DIM ft AS FILETIME
   IF pwszDateTime = NULL THEN RETURN ft
   DIM pDateTime AS Afx_ISWbemDateTime PTR = AfxNewCom("WbemScripting.SWbemDateTime")
   IF pDateTime = NULL THEN RETURN ft
   DIM bstrDateTime AS AFX_BSTR = SysAllocString(pwszDateTime)
   pDateTime->put_Value(bstrDateTime)
   IF bstrDateTime THEN SysFreeString(bstrDateTime)
   DIM bstrFileTime AS AFX_BSTR
   DIM hr AS HRESULT = pDateTime->GetFileTime(cast(VARIANT_BOOL, bIsLocal), @bstrFileTime)
   pDateTime->Release
   IF hr = &h80041001 THEN RETURN ft   ' // &h80041001 = wbemErrFailed
   DIM li AS LARGE_INTEGER
   li.QuadPart = _wtoi64(bstrFileTime)
   ft.dwLowDateTime = li.LowPart
   ft.dwHighDateTime = li.HighPart
   SysFreeString(bstrFileTime)
   RETURN ft

END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts a date and time value in the CIM DATETIME format to the FILETIME format.
' ========================================================================================
PRIVATE FUNCTION CWmiCli.WmiTimeToSystemTime (BYVAL pwszDateTime AS WSTRING PTR, BYVAL bIsLocal AS BOOLEAN = -1) AS SYSTEMTIME
   CWMICLI_DP("CWmiCli.WmiTimeToSystemTme")

   DIM ft AS FILETIME, st AS SYSTEMTIME
   ft = this.WmiTimeToFileTime(pwszDateTime, bIsLocal)
   FileTimeToSystemTime(@ft, @st)
   RETURN st

END FUNCTION
' ========================================================================================

END NAMESPACE
