' ########################################################################################
' Microsoft Windows
' File: CwmiDisp
' Contents: WMI classes
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2017 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "Afx/AfxWmiDisp.bi"
#include once "Afx/AfxCOM.inc"
#include once "Afx/CVar.inc"
#include once "Afx/CDispInvoke.inc"

NAMESPACE Afx

' ========================================================================================
' Returns the text string description associated with the WMI error code.
' ========================================================================================
PRIVATE FUNCTION AfxWmiGetErrorCodeText (BYVAL hRes AS HRESULT) AS CBSTR
   DIM cbsErrMsg AS CBSTR
   ' // Create an instance of the IWbemStatusCodeText interface
   DIM pIWbemStatusCodeText AS IWbemStatusCodeText PTR    ' // IWbemStatusCodeText interface
   DIM CLSID_WbemStatusCodeText_ AS CLSID = (&hEB87E1BD, &h3233, &h11D2, {&hAE, &hC9, &h00, &hC0, &h4F, &hB6, &h88, &h20})
   DIM IID_IWbemStatusCodeText_  AS GUID  = (&hEB87E1BC, &h3233, &h11D2, {&hAE, &hC9, &h00, &hC0, &h4F, &hB6, &h88, &h20})
   CoCreateInstance(@CLSID_WbemStatusCodeText_, NULL, CLSCTX_INPROC_SERVER, @IID_IWbemStatusCodeText_, @pIWbemStatusCodeText)
   IF pIWbemStatusCodeText THEN
      DIM hr AS HRESULT = pIWbemStatusCodeText->lpvtbl->GetErrorCodeText(pIWbemStatusCodeText, hRes, 0, 0, cbsErrMsg.vptr)
   END IF
   RETURN cbsErrMsg
END FUNCTION
' ========================================================================================

' ========================================================================================
' Helper function for IWbemStatusCodeText.GetFacilityCodeText.
' Returns the name of the subsystem where the error occurred, such as "Windows", "WBEM", "SSPI", or "RPC".
' ========================================================================================
PRIVATE FUNCTION AfxWmiGetFacilityCodeText (BYVAL hRes AS LONG) AS CBSTR
   DIM cbsErrMsg AS CBSTR
   ' // Create an instance of the IWbemStatusCodeText interface
   DIM pIWbemStatusCodeText AS IWbemStatusCodeText PTR    ' // IWbemStatusCodeText interface
   DIM CLSID_WbemStatusCodeText_ AS CLSID = (&hEB87E1BD, &h3233, &h11D2, {&hAE, &hC9, &h00, &hC0, &h4F, &hB6, &h88, &h20})
   DIM IID_IWbemStatusCodeText_  AS GUID  = (&hEB87E1BC, &h3233, &h11D2, {&hAE, &hC9, &h00, &hC0, &h4F, &hB6, &h88, &h20})
   CoCreateInstance(@CLSID_WbemStatusCodeText_, NULL, CLSCTX_INPROC_SERVER, @IID_IWbemStatusCodeText_, @pIWbemStatusCodeText)
   IF pIWbemStatusCodeText THEN
      DIM hr AS HRESULT = pIWbemStatusCodeText->lpvtbl->GetFacilityCodeText(pIWbemStatusCodeText, hRes, 0, 0, cbsErrMsg.vptr)
   END IF
   RETURN cbsErrMsg
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts a date and time value in the CIM DATETIME format to the FILETIME format.
' ========================================================================================
PRIVATE FUNCTION AfxWmiTimeToFileTime (BYVAL pwszDateTime AS WSTRING PTR, BYVAL bIsLocal AS BOOLEAN = TRUE) AS FILETIME
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
' Converts a date and time value in the CIM DATETIME format to a string containing the
' date based on the specified mask, e.g. "dd-MM-yyyy"
' ========================================================================================
PRIVATE FUNCTION AfxWmiTimeToDateStr (BYVAL pwszDateTime AS WSTRING PTR, BYREF wszMask AS WSTRING, BYVAL bIsLocal AS BOOLEAN = TRUE, BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS CBSTR
   DIM ft AS FILETIME, cbsDate AS CBSTR
   IF pwszDateTime = NULL THEN RETURN ""
   DIM pDateTime AS Afx_ISWbemDateTime PTR = AfxNewCom("WbemScripting.SWbemDateTime")
   IF pDateTime = NULL THEN RETURN ""
   DIM bstrDateTime AS AFX_BSTR = SysAllocString(pwszDateTime)
   pDateTime->put_Value(bstrDateTime)
   IF bstrDateTime THEN SysFreeString(bstrDateTime)
   DIM bstrFileTime AS AFX_BSTR
   DIM hr AS HRESULT = pDateTime->GetFileTime(cast(VARIANT_BOOL, bIsLocal), @bstrFileTime)
   pDateTime->Release
   IF hr = &h80041001 THEN RETURN ""   ' // &h80041001 = wbemErrFailed
   DIM li AS LARGE_INTEGER
   li.QuadPart = _wtoi64(bstrFileTime)
   ft.dwLowDateTime = li.LowPart
   ft.dwHighDateTime = li.HighPart
   SysFreeString(bstrFileTime)
   DIM ST AS SYSTEMTIME, wszDateStr AS WSTRING * 260
   FileTimeToSystemTime(@FT, @ST)
   GetDateFormatW(lcid, NULL, @ST, wszMask, wszDateStr, SIZEOF(wszDateStr))
   RETURN wszDateStr
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts a date and time value in the CIM DATETIME format to a string containing the
' the time based on the specified mask, e.g. "hh':'mm':'ss".
' ========================================================================================
PRIVATE FUNCTION AfxWmiTimeToTimeStr (BYVAL pwszDateTime AS WSTRING PTR, BYREF wszMask AS WSTRING, BYVAL bIsLocal AS BOOLEAN = TRUE, BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS CBSTR
   DIM ft AS FILETIME, cbsDate AS CBSTR
   IF pwszDateTime = NULL THEN RETURN ""
   DIM pDateTime AS Afx_ISWbemDateTime PTR = AfxNewCom("WbemScripting.SWbemDateTime")
   IF pDateTime = NULL THEN RETURN ""
   DIM bstrDateTime AS AFX_BSTR = SysAllocString(pwszDateTime)
   pDateTime->put_Value(bstrDateTime)
   IF bstrDateTime THEN SysFreeString(bstrDateTime)
   DIM bstrFileTime AS AFX_BSTR
   DIM hr AS HRESULT = pDateTime->GetFileTime(cast(VARIANT_BOOL, bIsLocal), @bstrFileTime)
   pDateTime->Release
   IF hr = &h80041001 THEN RETURN ""   ' // &h80041001 = wbemErrFailed
   DIM li AS LARGE_INTEGER
   li.QuadPart = _wtoi64(bstrFileTime)
   ft.dwLowDateTime = li.LowPart
   ft.dwHighDateTime = li.HighPart
   SysFreeString(bstrFileTime)
   DIM ST AS SYSTEMTIME, wszDateStr AS WSTRING * 260
   FileTimeToSystemTime(@FT, @ST)
   GetTimeFormatW(lcid, NULL, @ST, wszMask, wszDateStr, SIZEOF(wszDateStr))
   RETURN wszDateStr
END FUNCTION
' ========================================================================================

' ########################################################################################
' CWmiServices
' ########################################################################################
TYPE CWmiServices

Public:
   m_Result AS HRESULT
   m_pServices AS Afx_ISWbemServices PTR
   m_pObjectSet AS Afx_ISWbemObjectSet PTR
   m_pPropertySet AS Afx_ISWbemPropertySet PTR
   m_pEnum AS IEnumVARIANT PTR

Private:
   m_bUninitCOM AS BOOLEAN

Public:
   DECLARE CONSTRUCTOR (BYREF wszDisplayName AS WSTRING)
   DECLARE CONSTRUCTOR (BYREF cbsServer AS CBSTR, BYREF cbsNamespace AS CBSTR, _
      BYREF cbsUser AS CBSTR = "", BYREF cbsPassword AS CBSTR = "", BYREF cbsLocale AS CBSTR = "", _
      BYREF cbsAuthority AS CBSTR = "", BYVAL iSecurityFlags AS LONG = wbemConnectFlagUseMaxWait)
   DECLARE DESTRUCTOR
   DECLARE FUNCTION GetLastResult () AS HRESULT
   DECLARE FUNCTION SetResult (BYVAL Result AS HRESULT) AS HRESULT
   DECLARE FUNCTION GetErrorCodeText () AS CBSTR
   DECLARE FUNCTION GetFacilityCodeText () AS CBSTR
   DECLARE FUNCTION ServicesPtr () AS IDispatch PTR   ' Afx_ISWbemServices PTR
   DECLARE FUNCTION ServicesObj () AS IDispatch PTR   ' Afx_ISWbemServices PTR
   DECLARE FUNCTION ObjectSetPtr () AS IDispatch PTR   ' Afx_ISWbemObjectSet PTR
   DECLARE FUNCTION PropertySetPtr () AS IDispatch PTR   ' Afx_ISWbemPropertySet PTR
   DECLARE FUNCTION WmiTimeToFileTime (BYREF cbsDateTime AS CBSTR, BYVAL bIsLocal AS BOOLEAN = TRUE) AS FILETIME
   DECLARE FUNCTION WmiDateToStr (BYREF cbsDateTime AS CBSTR, BYREF wszMask AS WSTRING, BYVAL bIsLocal AS BOOLEAN = TRUE, BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS CBSTR
   DECLARE FUNCTION WmiDateToStr (BYREF cvDateTime AS CVAR, BYREF wszMask AS WSTRING, BYVAL bIsLocal AS BOOLEAN = TRUE, BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS CBSTR
   DECLARE FUNCTION WmiTimeToStr (BYREF cbsDateTime AS CBSTR, BYREF wszMask AS WSTRING, BYVAL bIsLocal AS BOOLEAN = TRUE, BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS CBSTR
   DECLARE FUNCTION WmiTimeToStr (BYREF cvDateTime AS CVAR, BYREF wszMask AS WSTRING, BYVAL bIsLocal AS BOOLEAN = TRUE, BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS CBSTR
   ' // Enumeration of the collection of objects generated by ExecQuery
   DECLARE FUNCTION ObjectsCount () AS LONG
   DECLARE FUNCTION NewEnum () AS BOOLEAN
   DECLARE FUNCTION ResetEnum () AS BOOLEAN
   DECLARE FUNCTION NextObject () AS CVAR
   ' // WMI methods
   DECLARE FUNCTION ExecQuery (BYREF cbsQuery AS CBSTR, BYVAL iFlags AS LONG = wbemFlagReturnWhenComplete) AS HRESULT
   DECLARE FUNCTION InstancesOf (BYREF cbsClass AS CBSTR, BYVAL iFlags AS LONG = wbemFlagReturnImmediately) AS HRESULT
'   DECLARE FUNCTION ExecMethod (BYREF cbsObjectPath AS CBSTR, BYREF cbsMethodName AS CBSTR, _
'           BYVAL objWbemInParameters AS Afx_IDispatch PTR, BYVAL iFlags AS LONG = 0, _
'           BYVAL objWbemNamedValueSet AS Afx_IDispatch PTR = NULL) AS Afx_ISWbemObject PTR
   DECLARE FUNCTION Get (BYREF cbsObjectPath AS CBSTR, BYVAL iFlags AS LONG = 0, BYVAL objWbemNamedValueSet AS Afx_IDispatch PTR = NULL) AS HRESULT
   ' // Properties
   DECLARE FUNCTION GetNamedProperties (BYVAL idx AS LONG = 0) AS HRESULT
   DECLARE FUNCTION PropsCount () AS LONG
   DECLARE PROPERTY PropValue (BYREF cbsName AS CBSTR) AS CVAR
   DECLARE PROPERTY PropValue (BYREF cbsName AS CBSTR, BYREF cvValue AS CVAR)

END TYPE
' ########################################################################################

' ========================================================================================
' Constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CWmiServices (BYREF wszDisplayName AS WSTRING)
   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
   ' // Connect to WMI using a moniker
   hr = CoGetObject(@wszDisplayName, NULL, @IID_IDispatch, cast(ANY PTR, @m_pServices))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Connects to the namespace that is specified on the cbsNamespace parameter on the computer
' that is specified in the cbsServer parameter. The target computer can be either local or
' remote, but it must have WMI installed.
' - cbsServer:
'   Computer name to which you are connecting. If the remote computer is in a different domain
'   than the user account under which you log in, then use the fully qualified computer name.
'   If you do not provide this parameter, the call defaults to the local computer.
'      Example: server1.network.fabrikam
'   You also can use an IP address in this parameter. If the IP address is in IPv6 format,
'   the target computer must be running IPv6. An address in IPv4 looks like 111.222.333.444
'   An IP address in IPv6 format looks like 2010:836B:4179::836B:4179
' - cbsNamespace:
'   String that specifies the namespace to which you log on. For example, to log on to the
'   root\default namespace, use root\default. If you do not specify this parameter, it
'   defaults to the namespace that is configured as the default namespace for scripting.
'      Example: DIM pServices AS CWmiServices = CWmiServices(".", "root\cimv2")
'      where "." is a shortcut for the local computer.
' - cbsUser [in, optional]
'   User name to use to connect. The string can be in the form of either a user name or a
'   Domain\Username. Leave this parameter blank to use the current security context. The
'   cbsUser parameter should only be used with connections to remote WMI servers. If you
'   attempt to specify cbsUser for a local WMI connection, the connection attempt fails.
'   If Kerberos authentication is in use, then the username and password that is specified
'   in cbsUser and cbsPassword cannot be intercepted on a network. You can use the UPN
'   format to specify the cbsUser.
'      Example: "DomainName\UserName"
'   Note: If a domain is specified in cbsAuthority, then the domain must not be specified
'   here. Specifying the domain in both parameters results in an Invalid Parameter error.
' - cbsPassword [in, optional]
'   String that specifies the password to use when attempting to connect. Leave the
'   parameter blank to use the current security context. The cbsPassword parameter should
'   only be used with connections to remote WMI servers. If you attempt to specify
'   cbsPassword for a local WMI connection, the connection attempt fails. If Kerberos
'   authentication is in use then the username and password that is specified in cbsUser
'   and cbsPassword cannot be intercepted on the network.
' - cbsLocale [in, optional]
'   String that specifies the localization code. If you want to use the current locale,
'   leave it blank. If not blank, this parameter must be a string that indicates the
'   desired locale where information must be retrieved. For Microsoft locale identifiers,
'   the format of the string is "MS_xxxx", where xxxx is a string in the hexadecimal form
'   that indicates the LCID. For example, American English would appear as "MS_409".
' - cbsAuthority [in, optional]
'   ""
'      This parameter is optional. However, if it is specified, only Kerberos or NTLMDomain
'      can be used.
'   Kerberos:
'      If the cbsAuthority parameter begins with the string "Kerberos:", then Kerberos
'      authentication is used and this parameter should contain a Kerberos principal name.
'      The Kerberos principal name is specified as Kerberos:domain, such as Kerberos:fabrikam
'      where fabrikam is the server to which you are attempting to connect.
'         Example: "Kerberos:DOMAIN"
'   NTLMDomain:
'      To use NT Lan Manager (NTLM) authentication, you must specify it as NTLMDomain:domain,
'      such as NTLMDomain:fabrikam where fabrikam is the name of the domain.
'         Example: "NTLMDomain:DOMAIN"
'      If you leave this parameter blank, the operating system negotiates with COM to
'      determine whether NTLM or Kerberos authentication is used. This parameter should
'      only be used with connections to remote WMI servers. If you attempt to set the
'      authority for a local WMI connection, the connection attempt fails.
'      Note: If the domain is specified in cbsUser, which is the preferred location, then
'     it must not be specified here. Specifying the domain in both parameters results in
'     an Invalid Parameter error.
' - iSecurityFlags [in, optional]
'     Used to pass flag values to ConnectServer.
'     0 (0x0)
'        A value of 0 for this parameter causes the call to ConnectServer to return only
'        after the connection to the server is established. This could cause your program
'        to stop responding indefinitely if the connection cannot be established.
'    wbemConnectFlagUseMaxWait (128 (0x80))
'        The ConnectServer call is guaranteed to return in 2 minutes or less. Use this flag
'        to prevent your program from ceasing to respond indefinitely if the connection
'        cannot be established.
' If successful, WMI returns an SWbemServices object that is bound to the namespace that
' is specified in cbsNamespace on the computer that is specified in cbsServer.
' Usage example with the local computer:
'    DIM pServices AS CWmiServices = CWmiServices(".", "root\cimv2")
'    IF pServices.ServicesPtr = NULL THEN END
' Remarks
'   The ConnectServer method is often used when connecting to an account with a different
'   username and password—credentials—on a remote computer because you cannot specify a
'   different password in a moniker string.
'   Using an IPv4 address to connect to a remote server may result in unexpected behavior.
'   The likely cause is stale DNS entries in your environment. In these circumstances, the
'   stale PTR entry for the machine will be used, with unpredictable results. To avoid
'   this behavior, you can append a period (".") to the IP address before calling
'   ConnectServer. This causes the reverse DNS lookup to fail, but may allow the
'   ConnectServer call to succeed on the correct machine.
' ========================================================================================
PRIVATE CONSTRUCTOR CWmiServices (BYREF cbsServer AS CBSTR, BYREF cbsNamespace AS CBSTR, _
   BYREF cbsUser AS CBSTR = "", BYREF cbsPassword AS CBSTR = "", BYREF cbsLocale AS CBSTR = "", _
   BYREF cbsAuthority AS CBSTR = "", BYVAL iSecurityFlags AS LONG = wbemConnectFlagUseMaxWait)
 
   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
   ' // Connect to the server
   DIM pLocator AS Afx_ISWbemLocator PTR = AfxNewCom("WbemScripting.SWbemLocator")
   IF pLocator THEN m_Result = pLocator->ConnectServer(cbsServer, cbsNamespace, cbsUser, cbsPassword, _
      cbsLocale, cbsAuthority, iSecurityFlags, NULL, @m_pServices)
   pLocator->Release

END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destructor
' ========================================================================================
PRIVATE DESTRUCTOR CWmiServices
   IF m_pServices THEN m_pServices->Release
   IF m_pObjectSet THEN m_pObjectSet->Release
   IF m_pEnum THEN m_pEnum->lpvtbl->Release(m_pEnum)
   IF m_pPropertySet THEN m_pPropertySet->Release
   IF m_bUninitCOM THEN CoUninitialize
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the last status code.
' ========================================================================================
PRIVATE FUNCTION CWmiServices.GetLastResult () AS HRESULT
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last status code.
' ========================================================================================
PRIVATE FUNCTION CWmiServices.SetResult (BYVAL Result AS HRESULT) AS HRESULT
   m_Result = Result
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the last error code description.
' ========================================================================================
PRIVATE FUNCTION CWmiServices.GetErrorCodeText () AS CBSTR
   RETURN AfxWmiGetErrorCodeText(m_Result)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns the name of the subsystem where the error occurred, such as "Windows", "WBEM", "SSPI", or "RPC".
' ========================================================================================
PRIVATE FUNCTION CWmiServices.GetFacilityCodeText () AS CBSTR
   RETURN AfxWmiGetFacilityCodeText(m_Result)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the underlying dispatch pointer.
' Don't call IUnknown_Release on it.
' ========================================================================================
PRIVATE FUNCTION CWmiServices.ServicesPtr () AS IDispatch PTR   ' Afx_ISWbemServices PTR
   RETURN cast(ANY PTR, m_pServices)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a counted reference of the underlying dispatch pointer.
' You must call IUnknown_Release when no longer need it.
' ========================================================================================
PRIVATE FUNCTION CWmiServices.ServicesObj () AS IDispatch PTR   ' Afx_ISWbemServices PTR
   IF m_pServices <> NULL THEN
      AfxSafeAddRef(m_pServices)
      RETURN cast(ANY PTR, m_pServices)
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a pointer to the ISWbemObjectSet interface.
' ========================================================================================
PRIVATE FUNCTION CWmiServices.ObjectSetPtr () AS IDispatch PTR   ' Afx_ISWbemObjectSet PTR
   RETURN cast(ANY PTR, m_pObjectSet)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a pointer to the ISWbemPropertySet interface.
' ========================================================================================
PRIVATE FUNCTION CWmiServices.PropertySetPtr () AS IDispatch PTR   ' Afx_ISWbemPropertySet PTR
   RETURN cast(ANY PTR, m_pPropertySet)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts a date and time value in the CIM DATETIME format to the FILETIME format.
' ========================================================================================
PRIVATE FUNCTION CWmiServices.WmiTimeToFileTime (BYREF cbsDateTime AS CBSTR, BYVAL bIsLocal AS BOOLEAN = TRUE) AS FILETIME
   RETURN AfxWmiTimeToFileTime(cbsDateTime, bIsLocal)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts a date and time value in the CIM DATETIME format to a string containing the
' date based on the specified mask, e.g. "dd-MM-yyyy"
' ========================================================================================
PRIVATE FUNCTION CWmiServices.WmiDateToStr (BYREF cbsDateTime AS CBSTR, BYREF wszMask AS WSTRING, BYVAL bIsLocal AS BOOLEAN = TRUE, BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS CBSTR
   RETURN AfxWmiTimeToDateStr(cbsDateTime, wszMask, bIsLocal, lcid)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CWmiServices.WmiDateToStr (BYREF cvDateTime AS CVAR, BYREF wszMask AS WSTRING, BYVAL bIsLocal AS BOOLEAN = TRUE, BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS CBSTR
   RETURN AfxWmiTimeToDateStr(cvDateTime.ToStr, wszMask, bIsLocal, lcid)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts a date and time value in the CIM DATETIME format to a string containing the
' the time based on the specified mask, e.g. "hh':'mm':'ss".
' ========================================================================================
PRIVATE FUNCTION CWmiServices.WmiTimeToStr (BYREF cbsDateTime AS CBSTR, BYREF wszMask AS WSTRING, BYVAL bIsLocal AS BOOLEAN = TRUE, BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS CBSTR
   RETURN AfxWmiTimeToTimeStr(cbsDateTime, wszMask, bIsLocal, lcid)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CWmiServices.WmiTimeToStr (BYREF cvDateTime AS CVAR, BYREF wszMask AS WSTRING, BYVAL bIsLocal AS BOOLEAN = TRUE, BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS CBSTR
   RETURN AfxWmiTimeToTimeStr(cvDateTime.ToStr, wszMask, bIsLocal, lcid)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the numbers of objects in the collection.
' Note: If ExecQuery has been called without the wbemFlagReturnWhenComplete flag, this
' method will return 0 because the operation has not been completed.
' ========================================================================================
PRIVATE FUNCTION CWmiServices.ObjectsCount () AS LONG
   DIM iCount AS LONG
   IF m_pObjectSet THEN m_Result = m_pObjectSet->get_Count(@iCount)
   RETURN iCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets a pointer to the enumerator of the collection of objects
' Note: If ExecQuery has been called using the wbemFlagReturnImmediately flag, the enumerator
' won't be available until the operation has completed.
' ========================================================================================
PRIVATE FUNCTION CWmiServices.NewEnum () AS BOOLEAN
   IF m_pEnum THEN AfxSafeRelease(m_pEnum)
   IF m_pObjectSet THEN m_Result = m_pObjectSet->get__NewEnum(cast(ANY PTR, @m_pEnum))
   IF m_pEnum THEN RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Resets the enumeration sequence to the beginning.
' ========================================================================================
PRIVATE FUNCTION CWmiServices.ResetEnum () AS BOOLEAN
   IF m_pEnum = NULL THEN RETURN FALSE
   FUNCTION = m_pEnum->lpvtbl->Reset(m_pEnum)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the value of the object
' ========================================================================================
PRIVATE FUNCTION CWmiServices.NextObject () AS CVAR
   DIM cvItem AS VARIANT
   IF m_pEnum = NULL THEN this.NewEnum
   DIM pceltFetched AS LONG
   IF m_pEnum THEN m_pEnum->lpvtbl->Next(m_pEnum, 1, @cvItem, @pceltFetched)
   RETURN cvItem
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves a named collection of the properties for the current class or instance.
' May return one of the error codes in the following list:
' wbemErrFailed - 2147749889 (0x80041001)
'    Unspecified failure.
' wbemErrInvalidParameter - 2147749896 (0x80041008)
'    Invalid parameter was specified.
' wbemErrOutOfMemory - 2147749896
'    Not enough memory for this method to execute.
' ========================================================================================
PRIVATE FUNCTION CWmiServices.GetNamedProperties (BYVAL idx AS LONG = 0) AS HRESULT
   IF m_pObjectSet = NULL THEN RETURN wbemErrFailed
   DIM pObject AS Afx_ISWbemObject PTR
   m_Result = m_pObjectSet->ItemIndex(idx, @pObject)
   IF m_pPropertySet THEN AfxSafeRelease(m_pPropertySet)
   IF pObject THEN pObject->get_Properties_(@m_pPropertySet)
   AfxSafeRelease(pObject)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Executes a query to retrieve objects. These objects are available through the returned
' SWbemObjectSet collection.
' ========================================================================================
PRIVATE FUNCTION CWmiServices.ExecQuery (BYREF cbsQuery AS CBSTR, BYVAL iFlags AS LONG = wbemFlagReturnWhenComplete) AS HRESULT
   IF m_pObjectSet THEN m_pObjectSet->Release
   m_Result = m_pServices->ExecQuery(cbsQuery, "WQL", iFlags, NULL, @m_pObjectSet)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates an enumerator that returns the instances of a specified class according to the
' user-specified selection criteria. This method implements a simple query. More complex
' queries may require the use of the ExecQuery method. By default, the method is called in
' the semisynchronous mode.
' ========================================================================================
PRIVATE FUNCTION CWmiServices.InstancesOf (BYREF cbsClass AS CBSTR, BYVAL iFlags AS LONG = wbemFlagReturnImmediately) AS HRESULT
   IF m_pObjectSet THEN m_pObjectSet->Release
   m_Result = m_pServices->InstancesOf(cbsClass, iFlags, NULL, @m_pObjectSet)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
'PRIVATE FUNCTION CWmiServices.ExecMethod (BYREF cbsObjectPath AS CBSTR, BYREF cbsMethodName AS CBSTR, _
'   BYVAL objWbemInParameters AS Afx_IDispatch PTR, BYVAL iFlags AS LONG = 0, _
'   BYVAL objWbemNamedValueSet AS Afx_IDispatch PTR = NULL) AS Afx_ISWbemObject PTR
'   DIM objWbemObject AS Afx_ISWbemObject PTR
'   m_Result = m_pServices->ExecMethod(cbsObjectPath, cbsMethodName, objWbemInParameters, iFlags, objWbemNamedValueSet, @objWbemObject)
'   RETURN objWbemObject
'END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves an object, that is either a class definition or an instance, based on the
' object path. This method retrieves only objects from the namespace that is associated
' with the current SWbemServices object.
' ========================================================================================
PRIVATE FUNCTION CWmiServices.Get (BYREF cbsObjectPath AS CBSTR, BYVAL iFlags AS LONG = 0, BYVAL objWbemNamedValueSet AS Afx_IDispatch PTR = NULL) AS HRESULT
   DIM pObject AS Afx_ISWbemObject PTR
   m_Result = m_pServices->Get(cbsObjectPath, iFlags, objWbemNamedValueSet, @pObject)
   IF m_Result <> S_OK THEN RETURN m_Result
   ' // Retrieve a reference to the ISWbemPropertySet interface (collection of properties)
   IF m_pPropertySet THEN m_pPropertySet->Release
   m_Result = pObject->get_Properties_(@m_pPropertySet)
   pObject->Release
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the numbers of objects in the property set collection.
' ========================================================================================
PRIVATE FUNCTION CWmiServices.PropsCount () AS LONG
   DIM iCount AS LONG
   IF m_pPropertySet THEN m_Result = m_pPropertySet->get_Count(@iCount)
   RETURN iCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the value of the named property.
' ========================================================================================
PRIVATE PROPERTY CWmiServices.PropValue (BYREF cbsName AS CBSTR) AS CVAR
   IF m_pPropertySet = NULL THEN EXIT PROPERTY
   DIM pProperty AS Afx_ISWbemProperty PTR
   m_Result = m_pPropertySet->Item(cbsName, 0, @pProperty)
   IF pProperty = NULL THEN EXIT PROPERTY
   DIM cvValue AS CVAR
   m_Result = pProperty->get_Value(cvValue.vptr)
   pProperty->Release
   PROPERTY = cvValue
END PROPERTY
' ========================================================================================

' ========================================================================================
' Sets the value of the property.
' ========================================================================================
PRIVATE PROPERTY CWmiServices.PropValue (BYREF cbsName AS CBSTR, BYREF cvValue AS CVAR)
   IF m_pPropertySet = NULL THEN EXIT PROPERTY
   DIM pProperty AS Afx_ISWbemProperty PTR
   m_Result = m_pPropertySet->Item(cbsName, 0, @pProperty)
   IF pProperty = NULL THEN EXIT PROPERTY
   m_Result = pProperty->put_Value(cvValue)
   pProperty->Release
END PROPERTY
' ========================================================================================

END NAMESPACE
